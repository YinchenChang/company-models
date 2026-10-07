# W1（2026-10-07）：產生 CRWV 欄位 ↔ Tokenomics 名稱對照表（給 W2）。
# 用法：python3 scripts/build_tk_map.py  → docs/plan/20261007_coreweave_Tokenomics對照表.xlsx 與同名 .md
# 數字全部讀 company.json 與 Tokenomics 快照（company.json → tokenomics.snapshotFile）；差距＝(CRWV − Tokenomics) ÷ Tokenomics。
# 比較世代（[Assumed]，W1 預設，W2 依世代組合改寫）：FY26–FY27 新增產能以 GB300、FY28–FY30 以 VR200 為代表世代。
import json, os
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CO = json.load(open(os.path.join(REPO, 'company.json'), encoding='utf-8'))
SN = json.load(open(os.path.join(REPO, CO['tokenomics']['snapshotFile']), encoding='utf-8'))
IT, SRC = SN['items'], SN['source']
D, M = CO['defaults'], CO['defaults']['m']
P = ['FY26 下半年', 'FY27', 'FY28', 'FY29', 'FY30']
REFG = ['GB300 NVL72', 'GB300 NVL72', 'VR200 NVL72', 'VR200 NVL72', 'VR200 NVL72']
GS = {'Hopper H100': 'H100', 'GB200 NVL72': 'GB200', 'GB300 NVL72': 'GB300', 'VR200 NVL72': 'VR200',
      'Rubin Ultra（MGX NVL 單架，72 GPU）': 'RU'}


def b(name, g):  # 基準值
    return IT[name]['values'][g]['基準']


def gens(name, scale=1.0, fmt='{:.2f}'):
    return '、'.join(f"{GS[g]} {fmt.format(v['基準'] * scale)}" for g, v in IT[name]['values'].items())


def gap(x, y):
    return None if (x is None or y in (None, 0)) else x / y - 1


def fl(xs, f='{:g}'):
    return '、'.join(f.format(x) for x in xs)


# 站點市場基準（segA siteBenchQ 同一算法）：具名站點 合約總額 ÷ 年數 ÷ 規劃 MW
_s = [s for s in CO['sites'] if not s.get('residual') and s.get('contract') and s.get('years') and s.get('planned')] \
    if 'sites' in CO else [s for s in D['sites'] if not s.get('residual') and s.get('contract') and s.get('years') and s.get('planned')]
BENCH = sum(s['contract'] / s['years'] for s in _s) / sum(s['planned'] for s in _s) * 1e3
PWR = [8760 * pu * pw / 1e6 for pu, pw in zip(M['pue'], M['power'])]  # $M/MW/年（segA：MW×8760×PUE×$/MWh）
TK_PWR = IT['IF_PowerCost']['values']['GB300 NVL72']['基準']  # $B/GW/年＝$M/MW/年（與世代無關）
HE = [b('IF_HoldEcon', g) for g in REFG]
CX_IT = [b('IF_CapexIT', g) for g in REFG]
CX_TOT = [b('IF_CapexTotal', g) for g in REFG]
REV = [x * 1e3 for x in M['revMW']]
REF_TXT = 'FY26–27＝GB300、FY28–30＝VR200（[Assumed]，W1 比較用）'
ROWS = [
    # (CRWV 欄位, Excel 列, 現值, 單位與 MW 口徑, 現值依據, Tokenomics 名稱, Tokenomics 基準（CRWV 單位）, 差距（各期）, W2 處理, 備註)
    ('scenarios.capexTemplate.costMW', '輸入與假設｜每 MW 建置成本', fl(CO['scenarios']['capexTemplate']['costMW']),
     'US$m/MW；CRWV MW（視為 IT 關鍵電力，見 MW 口徑）',
     'README：「每 MW 建置成本（GPU＋網路＋機房內裝），各期」；Excel 說明：「以 FY26 校準重現全年指引 35–39 [Derived]」（＝以 CapEx 指引回推，共同規則第 4 節不再允許作為參數來源）',
     'IF_CapexIT（＋自建部分的 IF_CapexFacility）', fl(CX_IT, '{:.2f}') + f'（IF_CapexIT，{REF_TXT}）；各世代：' + gens('IF_CapexIT'),
     [gap(x, y) for x, y in zip(CO['scenarios']['capexTemplate']['costMW'], CX_IT)], '引用',
     'CRWV 機房多為租賃（廠房以租金計），所以對應 IT 設備而非合計；IF_CapexTotal 各世代：' + gens('IF_CapexTotal') + '。若 MW 為設施口徑，CRWV 34 換成每 IT MW＝34×1.2＝40.8'),
    ('defaults.gpuLife', '輸入與假設｜GPU 經濟壽命（年）', str(D['gpuLife']), '年',
     'Excel 說明：「公司對伺服器與網路設備的折舊年限；決定汰換時點與車隊折舊 [Verified]」',
     'IF_DeprLifeIT（v5.25）', f"{SRC['version']} 無此名稱（快照 missing）", [None] * 5, '引用（v5.25 後）',
     'W2 第 0 步確認 Tokenomics master 是否已有 v5.25；沒有則維持 6 年並列為未解'),
    ('defaults.m.power ＋ defaults.m.pue', '輸入與假設｜電價、PUE（overlay 用）', f"電價 {fl(M['power'])} $/MWh；PUE {fl(M['pue'])}；合成電費 " + fl(PWR, '{:.3f}') + ' US$m/MW/年',
     'US$m/MW/年（CRWV：Accepted MW × 8760 × PUE × 電價，即 100% 滿載）',
     'Excel 說明：「overlay 用 [Assumed]」；segA：accepted × 8760 × pue × power',
     'IF_PowerCost（另 IF_FacilityGW、v5.25 IF_PowerPrice、IF_AvgDraw）', f'{TK_PWR:.3f}（每 IT MW；與世代無關）',
     [gap(x, TK_PWR) for x in PWR], '引用',
     f"IF_FacilityGW 基準 {b('IF_FacilityGW', 'GB300 NVL72'):g}（PUE 對照：CRWV {M['pue'][0]}–{M['pue'][-1]}，差距 {gap(M['pue'][0], 1.2):+.1%}～{gap(M['pue'][-1], 1.2):+.1%}）；電價與平均用電係數待 v5.25 拆開"),
    ('defaults.m.maint', '輸入與假設｜維護成本（overlay 用）', fl(M['maint']), 'US$m/MW/年',
     'Excel 說明：「overlay 用 [Assumed]」', 'IF_MaintIT＋IF_MaintFac（v5.25）', f"{SRC['version']} 無此名稱（快照 missing）",
     [None] * 5, '引用（v5.25 後）', '另 IF_StaffSW、IF_TaxIns（v5.25）可補齊由下而上營運費用'),
    ('defaults.m.revMW', '輸入與假設｜每 MW 年收入', fl(M['revMW'], '{:.4f}') + ' US$bn（＝' + fl(REV, '{:.1f}') + ' US$m）', 'US$bn/MW/年；Billable MW',
     'Excel 說明：「期末 ARR 18.5–19.5÷1.85GW≈10.0–10.5M/MW；7 月漲價 25% 只影響新約 [Derived]」（＝以 ARR 回推，共同規則第 4 節不再允許）',
     'IF_HoldEcon（下限對照）；W2 改為 GPU 小時價格 × IF_GPUsPerGW × 8760 × 利用率', fl(HE, '{:.2f}') + f' US$m/MW/年（IF_HoldEcon 經濟持有成本，{REF_TXT}）；各世代：' + gens('IF_HoldEcon'),
     [gap(x, y) for x, y in zip(REV, HE)], '推導',
     '經濟持有成本＝100% 時數、不含利潤的「打平收入」；CRWV 每 MW 收入低於 GB300／VR200 的持有成本即表示新世代每 MW 付不起資本成本。L1_RevGW_Fleet_VR200 ' + f"{IT['L1_RevGW_Fleet_VR200']['values']:.1f}" + ' 為 token 層理想上限（OpenAI 單價），與 GPU 租賃收入不同層，只作上限對照'),
    ('defaults.ebStart／defaults.ebSteady', '輸入與假設｜起始／穩態 EBITDA 率', f"{D['ebStart']:.0%} → {D['ebSteady']:.0%}（線性）", '比例',
     '交接檔：「由下而上（每 MW 現金成本約 $3.2–3.6m／年收入約 $10.9m）約 67–71%；取 65%」；起始＝Q2 實際 Adj. EBITDA 率 58.6% [Verified]',
     'IF_OpexGW（v5.25；EBITDA 率＝1 − 營運費用 ÷ 每 MW 收入 − 管銷率）', f"{SRC['version']} 無 IF_OpexGW；現有可得部分：電費 IF_PowerCost {TK_PWR:.3f} ÷ 每 MW 收入 {REV[0]:.1f} ＝ {TK_PWR / REV[0]:.1%} 營收",
     [None] * 5, '推導（v5.25 後）', 'W2 依工作單以由下而上成本＋管銷率推導；管銷率見 permw_inputs（10-Q 實際）'),
    ('defaults.m.util', '輸入與假設｜利用率', fl(M['util']) + ' %', '%（已售出 ÷ Billable MW）',
     'Excel 說明：「法說「近期產能實質售罄」，不擬合 100% [Assumed]」', 'IF_Util（口徑不同）',
     f"{IT['IF_Util']['values'] * 100:.0f}%（Tokenomics token 服務基準利用率）", [gap(x / 100, IT['IF_Util']['values']) for x in M['util']], '不動（公司專屬）',
     '口徑不同：CRWV＝長約售出比例（take-or-pay 計費），Tokenomics＝token 服務的算力利用率；差距不代表衝突，W2 不引用 IF_Util'),
    ('站點每 MW 年租金市場基準（methodology.checks.rentVsBenchMin 所用）', '運營_站點｜加權每 MW 年租金（H 欄）', f'{BENCH:.3f}（檢驗門檻＝× 第三方占比 {CO["leases"]["facts"]["share"]} × {CO["methodology"]["checks"]["rentVsBenchMin"]}）',
     'US$m/MW/年；具名站點 critical IT MW（Helios、Polaris Forge 1、Core Scientific）',
     'segA siteBenchQ：具名站點 合約額 ÷ 年數 ÷ 規劃 MW；交接檔：「模型每 MW 年租金 FY30 約 $0.96m，低於具名站點基準 $1.70m × 85%」',
     'Tokenomics 無對應（公司專屬）；對照 L1_FacCapexMW', f"L1_FacCapexMW {IT['L1_FacCapexMW']['values']:.2f} US$m/MW（廠房資本支出）；租金 ÷ 廠房造價＝{BENCH / IT['L1_FacCapexMW']['values']:.1%}／年",
     [None] * 5, '對照', '租金是房東的廠房資本回收：13% 左右的年租金收益率對 12.67 的廠房造價屬合理區間（W2 只作對照，不改租金來源）'),
    ('methodology.checks.capexPerMwBand', '檢查_連動｜CapEx 強度（模型期合計）', fl(CO['methodology']['checks']['capexPerMwBand']), 'US$m/MW',
     '檢查頁說明：「公司自身比率：2026 CapEx 35–39bn 對應約 +1.25 GW ＝ $28–34m/MW」',
     'IF_CapexIT（下端）／IF_CapexTotal（上端）', f"IF_CapexIT 基準 {min(v['基準'] for v in IT['IF_CapexIT']['values'].values()):.2f}–{max(v['基準'] for v in IT['IF_CapexIT']['values'].values()):.2f}；IF_CapexTotal 基準 {min(v['基準'] for v in IT['IF_CapexTotal']['values'].values()):.2f}–{max(v['基準'] for v in IT['IF_CapexTotal']['values'].values()):.2f}；低成本最小 {min(v['低成本'] for v in IT['IF_CapexIT']['values'].values()):.2f}、高成本最大 {max(v['高成本'] for v in IT['IF_CapexTotal']['values'].values()):.2f}",
     [gap(CO['methodology']['checks']['capexPerMwBand'][0], min(v['低成本'] for v in IT['IF_CapexIT']['values'].values())),
      gap(CO['methodology']['checks']['capexPerMwBand'][1], b('IF_CapexTotal', 'GB300 NVL72'))], '推導',
     '差距兩格＝下端對 IF_CapexIT 低成本最小值（GB200）、上端對 IF_CapexTotal GB300 基準；W2 建議改為「代表世代 IT 低成本～合計高成本」'),
    ('（CRWV 無此欄位）每 MW GPU 數', '—（W2 新增）', '—', '顆/MW（IT 關鍵電力）', 'CRWV 目前不建 GPU 數；W2 GPU 小時價格路線需要',
     'IF_GPUsPerGW', gens('IF_GPUsPerGW', 1e-3, '{:.0f}') + ' 顆/IT MW', [None] * 5, '引用',
     'S-1 年底對照：約 250,000 GPU ÷ 360 MW ≈ 694 顆/MW（[Derived]，只作 MW 口徑對照；Hopper 為主）'),
    ('（CRWV 無此欄位）每 GPU 小時持有成本', '—（W2 新增）', '—', 'US$/GPU-hr', 'CRWV 目前不建；W2 GPU 小時價格路線的成本下限',
     'IF_GPUhrEcon；L1_GPUhr_GB200_vsCW、L1_GPUhr_GB300_vsBE', gens('IF_GPUhrEcon'), [None] * 5, '對照',
     f"L1 外部對照：GB200 對 CoreWeave 隨需牌價 {IT['L1_GPUhr_GB200_vsCW'].get('external', {}).get('外部低')}（{IT['L1_GPUhr_GB200_vsCW'].get('external', {}).get('srcId')}）；GB300 對新雲損益兩平租金 {IT['L1_GPUhr_GB300_vsBE'].get('external', {}).get('外部低')}（{IT['L1_GPUhr_GB300_vsBE'].get('external', {}).get('srcId')}）"),
]

HDR = ['CRWV 欄位（company.json 路徑）', 'Excel 列名', '現值（各期：' + '、'.join(P) + '）', '單位與 MW 口徑', '現值依據（README／交接檔原文）',
       'Tokenomics 名稱', f"Tokenomics 基準值（換成 CRWV 單位；{SRC['version']}）", '差距 %（CRWV ÷ Tokenomics − 1）', 'W2 處理', '備註']


def gtxt(g):
    xs = [x for x in g if x is not None]
    if not xs: return '—（無可比值）'
    return '、'.join('—' if x is None else f'{x:+.1%}' for x in g)


def main():
    out_dir = os.path.join(REPO, 'docs', 'plan')
    os.makedirs(out_dir, exist_ok=True)
    base = os.path.join(out_dir, '20261007_coreweave_Tokenomics對照表')
    wb = Workbook(); ws = wb.active; ws.title = '對照表'
    ws['A1'] = 'CRWV 欄位 ↔ Tokenomics 名稱對照表（W1，給 W2）'; ws['A1'].font = Font(bold=True, size=13)
    ws['A2'] = (f"Tokenomics {SRC['version']}（{SRC['file']}，commit {SRC['commit'][:7]}，擷取 {SRC['extractedAt']}）；基準成本情境；每 GW＝IT 關鍵電力。"
                f"差距＝CRWV ÷ Tokenomics − 1；比較世代 {REF_TXT}。CRWV MW 口徑：視為 IT 關鍵電力 [Assumed]（見 permw_inputs_20261007.json）。")
    for j, h in enumerate(HDR):
        c = ws.cell(row=4, column=1 + j, value=h); c.font = Font(bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor='1F3864')
        c.alignment = Alignment(wrap_text=True, vertical='top')
    for i, row in enumerate(ROWS):
        vals = list(row[:7]) + [gtxt(row[7])] + list(row[8:])
        for j, v in enumerate(vals):
            c = ws.cell(row=5 + i, column=1 + j, value=v); c.alignment = Alignment(wrap_text=True, vertical='top')
    for col, w in zip('ABCDEFGHIJ', (30, 26, 30, 24, 50, 30, 46, 22, 14, 50)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = 'B5'
    wb.save(base + '.xlsx')
    md = [f"# CRWV 欄位 ↔ Tokenomics 名稱對照表（W1，給 W2）", '',
          f"- Tokenomics {SRC['version']}（`{SRC['file']}`，commit `{SRC['commit'][:7]}`，擷取 {SRC['extractedAt']}）；基準成本情境；每 GW＝IT 關鍵電力。",
          f"- 差距＝CRWV ÷ Tokenomics − 1；比較世代 {REF_TXT}。CRWV MW 口徑：視為 IT 關鍵電力 [Assumed]（依據見 `coreweave/data/permw_inputs_20261007.json`）。",
          '- 同內容 Excel：`20261007_coreweave_Tokenomics對照表.xlsx`（由 `coreweave/scripts/build_tk_map.py` 產生）。', '',
          '| ' + ' | '.join(HDR) + ' |', '|' + '---|' * len(HDR)]
    for row in ROWS:
        vals = list(row[:7]) + [gtxt(row[7])] + list(row[8:])
        md.append('| ' + ' | '.join(str(v).replace('|', '／').replace('\n', ' ') for v in vals) + ' |')
    open(base + '.md', 'w', encoding='utf-8').write('\n'.join(md) + '\n')
    print('saved', base + '.xlsx', base + '.md')
    for row in ROWS:
        print(row[0][:40], '|', gtxt(row[7]))


if __name__ == '__main__':
    main()
