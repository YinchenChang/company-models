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
if CO.get('tokenomics'):  # W6 r1：k 基準與區間依目前快照推算（與 build_xlsx、HTML 同一規則：kspec.py）
    sys.path.insert(0, ROOT); import kspec
    kspec.resolve(CO, json.load(open(os.path.join(ROOT, CO['tokenomics']['snapshotFile']), encoding='utf-8'))['items'])
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
if PMW['revenue'] == 'tkAnchor':  # W4 r2：定價倍數 k 與隨需占比（與 HTML pmwSensQ 同一組設定、同一名稱）
    _AM = CO['pricing']['anchorMultiple']; _KL, _OD = '定價倍數 k_長約（市場長約價 ÷ Tokenomics 同世代持有成本）', '隨需占比（占在役計費產能）'
    CASES += [('kLongLo', f"k_長約 {_AM['long']['low']:.2f}", {_KL: _AM['long']['low']}), ('kLongMed', f"k_長約 {_AM['long']['sensMedian']:.2f}（三筆長約中位數）", {_KL: _AM['long']['sensMedian']}),
              ('kLongHi', f"k_長約 {_AM['long']['high']:.2f}", {_KL: _AM['long']['high']})]
    CASES += [(f'od{j + 1}', f"隨需占比 {round(x * 100)}%（k_現貨 {_AM['spot']['base']:.2f}）", {_OD: x}) for j, x in enumerate(_AM['onDemandShare']['sens'])]
if not CO.get('fleet'):
    print('permw_sens：company.json 無 fleet，不適用'); sys.exit(0)
_CA = CO.get('companyAdjust') if PMW['revenue'] == 'tkAnchor' else None
sys.path.insert(0, ROOT); import calendar_q as _cq; _YTDL = _cq.derive(CO)['ytdLabel']  # W5：年初至今標籤（滾動後自動更新）
with Workbook(xlsx) as wb:
    if _CA:  # W5：公司實況驗證的敏感度（輸入值讀 Excel「公司實況驗證」頁敏感度輸入列；與 HTML pmwSensQ／cvSensQ 同一組設定、同一名稱）
        _CV = '公司實況驗證'; _g = lambda lab: wb.get(wb.cell(_CV, lab))
        _z, _zc, _kc = _g('敏感度輸入｜營運成本倍數＝Q2 實際 ÷ Q2 季末世代 Tokenomics 合計'), _g(f"敏感度輸入｜每 MW 建置成本倍數＝{_YTDL} 實際 ÷ 首期新增世代 Tokenomics"), _g('敏感度輸入｜CRWV 短天期 k（短天期合約價 ÷ 同世代 Tokenomics，基準）')
        _OX, _CXS, _KS = '由下而上營運成本倍數（1＝Tokenomics 值）', '每 MW 建置成本倍數（整體）', '定價倍數 k_現貨（市場現貨價 ÷ Tokenomics 同世代持有成本）'
        CASES += [('kExOff', '既有合約 k 不套用（全部按 k_新約）', {'既有合約 k 開關（1＝套用、0＝不套用）': 0}),
                  ('kNewUp', f"新約價格 +{round(_CA['newK']['adjSens'] * 100)}%（公司說法，只作用於長約）", {'新約價格調整（相對 k_長約；只作用於長約部分）': _CA['newK']['adjSens']}),
                  ('spotCw', f"隨需 {round(_CA['spotCw']['od'] * 100)}% × CRWV 短天期 k {_kc:.2f}", {_OD: _CA['spotCw']['od'], _KS: _kc}),
                  ('opexQ2', f"營運成本＝Q2 實際比率（× {_z:.2f}）", {_OX: _z}),
                  ('capexQ2', f"每 MW 建置成本＝{_YTDL} 實際比率（× {_zc:.2f}）", {_CXS: _zc}),
                  ('actBoth', '營運成本與建置成本皆用公司實際比率', {_OX: _z, _CXS: _zc})]
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
