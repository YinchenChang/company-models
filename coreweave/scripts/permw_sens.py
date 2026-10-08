# W2（2026-10-07）：每 MW 經濟性敏感度的建置時快照——以 LibreOffice（UNO，scripts/uno_q.py）開啟建好的 Excel，
# 依序切換「情境選擇」（保守／基準／積極）與敏感度輸入（Tokenomics 成本情境、GPU 小時價格情境、世代組合 Rubin Ultra 版、管銷率 GAAP 口徑），
# 讀重算後的「加權目標價」與融資缺口（＝MAX(0, −FY30 融資前累積現金)），寫入 permw_sens.json；build_xlsx.py 讀此檔寫入「每MW經濟性」頁「敏感度」區，
# 並以「快照狀態」格比對目前輸入下的加權目標價（不一致＝快照已過期）。HTML 以同一組設定即時計算（pmwSensQ），cmp31 逐格比對。
# 用法：python3 scripts/permw_sens.py 檔案.xlsx [輸出 json，預設 repo 根目錄 permw_sens.json]；結果與原檔相同時代碼 0，有變動時寫入並以代碼 3 結束。
import sys, os, json
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from uno_q import Workbook

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
xlsx = sys.argv[1]; out = sys.argv[2] if len(sys.argv) > 2 else os.path.join(ROOT, 'permw_sens.json')
CO = json.load(open(os.path.join(ROOT, 'company.json'), encoding='utf-8'))
PMW = {'capex': 'legacy', 'cost': 'ebitdaPct', 'revenue': 'legacy', **(CO.get('methodology', {}).get('perMw') or {})}
IN, VA, FR = '輸入與假設', '評價_DCF與目標價', '各期收支'
# (代碼, 名稱, {輸入列名稱: 值}；None＝本方法不適用)
CASES = [('base', '基準（目前輸入）', {}),
         ('tkLow', 'Tokenomics 低成本', {'Tokenomics 成本情境（1＝低成本、2＝基準、3＝高成本）': 1}),
         ('tkHigh', 'Tokenomics 高成本', {'Tokenomics 成本情境（1＝低成本、2＝基準、3＝高成本）': 3}),
         ('pxLow', 'GPU 小時價格 低', {'GPU 小時價格情境（1＝低、2＝基準、3＝高）': 1} if PMW['revenue'] == 'gpuHr' else None),
         ('pxHigh', 'GPU 小時價格 高', {'GPU 小時價格情境（1＝低、2＝基準、3＝高）': 3} if PMW['revenue'] == 'gpuHr' else None),
         ('mixRU', '世代組合 Rubin Ultra 版', {'世代組合敏感度（0＝基準、1＝Rubin Ultra 版）': 1}),
         ('sgaGaap', '管銷率 GAAP（含 SBC）', {'管銷率口徑（1＝扣 SBC、2＝GAAP 含 SBC）': 2} if PMW['cost'] == 'bottomUp' else None)]
if PMW['revenue'] == 'tkAnchor':  # W4：定價倍數 k 與長約占比（與 HTML pmwSensQ 同一組設定、同一名稱）
    _AM = CO['pricing']['anchorMultiple']; _KL, _KS, _LS = '定價倍數 k_長約（市場長約價 ÷ Tokenomics 同世代持有成本）', '定價倍數 k_現貨（市場現貨價 ÷ Tokenomics 同世代持有成本）', '長約占比調整（百分點；0＝RPO 涵蓋估計）'
    CASES += [('kLongLo', f"k_長約 {_AM['long']['low']:.2f}", {_KL: _AM['long']['low']}), ('kLongHi', f"k_長約 {_AM['long']['high']:.2f}", {_KL: _AM['long']['high']}),
              ('kSpotLo', f"k_現貨 {_AM['spot']['low']:.2f}", {_KS: _AM['spot']['low']}), ('kSpotHi', f"k_現貨 {_AM['spot']['high']:.2f}", {_KS: _AM['spot']['high']}),
              ('lsLow', f"長約占比 {round(_AM['longShare']['sensLowPt'] * 100)}pt", {_LS: _AM['longShare']['sensLowPt']}), ('lsAll', '長約占比 100%（新簽約全視為長約）', {_LS: 1})]
if not CO.get('fleet'):
    print('permw_sens：company.json 無 fleet，不適用'); sys.exit(0)
with Workbook(xlsx) as wb:
    c_sel = wb.cell(IN, '情境選擇', prefix=True)
    o_tgt, o_pfc = wb.cell(VA, '加權目標價'), wb.cell(FR, '融資前累積現金', col='G')
    cells = {k: wb.cell(IN, k) for _, _, ch in CASES if ch for k in ch}
    base_in = {k: wb.get(c) for k, c in cells.items()}; sel0 = wb.get(c_sel)
    res = {'cases': [[k, n] for k, n, _ in CASES], 'scenarios': {}}
    for s, sk in ((1, 'low'), (2, 'base'), (3, 'high')):
        wb.set(c_sel, s); row = {}
        for k, n, ch in CASES:
            if ch is None: row[k] = None; continue
            for kk, c in cells.items(): wb.set(c, ch.get(kk, base_in[kk]))
            row[k] = [wb.get(o_tgt), max(0.0, -wb.get(o_pfc))]
        for kk, c in cells.items(): wb.set(c, base_in[kk])
        res['scenarios'][sk] = row
    wb.set(c_sel, sel0)
num = lambda x: int(x) if isinstance(x, float) and x.is_integer() else x
res['scenarios'] = {s: {k: (None if v is None else [num(x) for x in v]) for k, v in row.items()} for s, row in res['scenarios'].items()}
txt = json.dumps(res, separators=(',', ':'), ensure_ascii=False)
old = open(out, encoding='utf-8').read() if os.path.exists(out) else None
b = res['scenarios']['base']
print(f"每 MW 敏感度快照（Excel 求值）：基準情境加權目標價 {b['base'][0]:.2f}、Tokenomics 低／高 {b['tkLow'][0]:.2f}／{b['tkHigh'][0]:.2f}、Rubin Ultra 版 {b['mixRU'][0]:.2f}")
if old == txt:
    print('permw_sens.json 無變動'); sys.exit(0)
open(out, 'w', encoding='utf-8').write(txt)
print(f'permw_sens.json 已更新（{out}）'); sys.exit(3)
