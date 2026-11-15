# 估值頁（Val）建置快照（v0.2.3）：情境 A／C 與敏感度（龍捲風）。
# 做法同 scripts/sensitivity.py：以 build_xlsx.py --set 改寫 Inputs → LibreOffice 重算 → 讀 Val 頁值；Python 不做模型計算。
# 用法（在 minimax/ 內）：
#   python3 scripts/val_snapshots.py scen   → data/val_scenarios.json（情境 A／B／C 的 Val 第 3 節來源列；B 只供過期檢查）
#   python3 scripts/val_snapshots.py sens   → data/val_sens.json（各案例低／高值的加權目標價；須先有 val_scenarios.json）
# 之後重跑 build_xlsx.py 才會把快照寫入 Val 頁（scripts/build_all.sh 已依序執行）。
import datetime as dt
import json
import os
import subprocess
import sys
from concurrent.futures import ThreadPoolExecutor

from openpyxl import load_workbook

HERE = os.path.dirname(os.path.abspath(__file__)); ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
from recalc import recalc  # noqa: E402

TOK_HI = ['I09@2027=4.0', 'I09@2028=1.5', 'I09@2029=0.9', 'I09@2030=0.6']
TOK_LO = ['I09@2027=0.8', 'I09@2028=0.5', 'I09@2029=0.3', 'I09@2030=0.2']
SCN = {'A': TOK_HI, 'B': [], 'C': TOK_LO}
SRCK = ('rev', 'ebit', 'cx', 'ma')
# (鍵, 名稱, 白話, 低值設定, 高值設定, 低值說明, 高值說明, 低值白話, 高值白話)
CASES = [
    ('wacc', 'WACC ±1pt（I40）', '折現率（WACC）', ['I40=-0.01'], ['I40=0.01'], 'WACC −1pt', 'WACC +1pt', '折現率低 1 個百分點', '折現率高 1 個百分點'),
    ('g', '永續成長率 2%／4%（I41）', '永續成長率', ['I41=0.02'], ['I41=0.04'], 'g＝2%', 'g＝4%', '永續成長 2%', '永續成長 4%'),
    ('beta', 'β 1.0／1.6（I36）', '股價波動係數（β）', ['I36=1.0'], ['I36=1.6'], 'β＝1.0', 'β＝1.6', '波動係數 1.0', '波動係數 1.6'),
    ('erp', 'ERP 4%／6%（I37）', '股權風險溢酬', ['I37=0.04'], ['I37=0.06'], 'ERP＝4%', 'ERP＝6%', '風險溢酬 4%', '風險溢酬 6%'),
    ('extg', '延伸期終點營收年增 5%／15%（I45）', '2031–2033 營收成長終點', ['I45=0.05'], ['I45=0.15'], '5%', '15%', '2033 年營收年增 5%', '2033 年營收年增 15%'),
    ('extm', '延伸期營業利益率加碼 0／+110pt（I46）', '2031–2033 營業利益率改善', ['I46=0'], ['I46=1.1'], '維持 FY2030', '+110pt（2033 年約 +13%）', '利益率維持 2030 年水準', '2033 年利益率轉為約 +13%'),
    ('wdcf', 'DCF 權重 25%／75%（I52）', 'DCF 與同業法的權重', ['I52=0.25'], ['I52=0.75'], 'DCF 25%', 'DCF 75%', 'DCF 占 25%', 'DCF 占 75%'),
    ('tax', '稅率 15%／25%（I39）', '所得稅率', ['I39=0.15'], ['I39=0.25'], '15%', '25%', '稅率 15%', '稅率 25%'),
    ('wc', '營運資金率 −5%／+5%（I48）', '營運資金占營收增量', ['I48=-0.05'], ['I48=0.05'], '−5%', '+5%', '預收款增加（−5%）', '營運資金占用（+5%）'),
    ('adj', '相對倍數調整：不調整／依 FY2030 利益率調整（I50）', '同業倍數是否依獲利差異打折', ['I50=0'], ['I50=1'], '不調整', '調整', '不打折', '依 2030 年虧損程度打折'),
    ('tok', 'API token 量年增低端／高端（I09；情境 A／C 的驅動）', 'API token 量成長', TOK_LO, TOK_HI, '低端', '高端', 'token 量成長慢', 'token 量成長快'),
    ('px', 'API 單價每年 −50%／持平（I08）', 'API 單價年變動', ['I08=-0.5'], ['I08=0'], '−50%／年', '持平', '單價每年腰斬', '單價不再下跌'),
    ('train', '訓練 MW 成長 ×0.5／×1.5（I17）', '訓練算力擴張速度', ['I17=*1.5'], ['I17=*0.5'], '×1.5', '×0.5', '訓練算力擴張快 1.5 倍', '訓練算力擴張減半'),
]


def run(tag, sets):
    d = os.path.join(ROOT, 'out', 'val'); os.makedirs(d, exist_ok=True)
    f = os.path.join(d, f'{tag}.xlsx')
    args = ['python3', os.path.join(ROOT, 'build_xlsx.py'), f]
    for s in sets:
        args += ['--set', s]
    subprocess.run(args, check=True, capture_output=True)
    r = recalc(f)
    if r['status'] != 'success':
        raise SystemExit(f'{tag}: {r}')
    return f


def labels(ws):
    return {ws.cell(r, 1).value: r for r in range(1, ws.max_row + 1) if isinstance(ws.cell(r, 1).value, str)}


def read_src(f):
    ws = load_workbook(f, data_only=True)['Val']
    lab = labels(ws)
    out = {}
    for k in SRCK:
        r = next(r for t, r in lab.items() if t.endswith('（B）') and r > lab[next(x for x in lab if x.startswith('情境 B：'))] and
                 t.startswith({'rev': '總營收', 'ebit': 'EBIT', 'cx': '資本支出', 'ma': '併購'}[k]))
        out[k] = [float(ws.cell(r, c).value) for c in range(3, 8)]
    r = next(r for t, r in lab.items() if t.startswith('營收：FY2025') and t.find('（B；') > 0)
    out['fy'] = [float(ws.cell(r, c).value) for c in range(3, 6)]
    return out


def read_tp(f):
    ws = load_workbook(f, data_only=True)['Val']
    r = next(r for t, r in labels(ws).items() if t.startswith('加權目標價＝'))
    return float(ws.cell(r, 3).value)


def scen():
    with ThreadPoolExecutor(3) as ex:
        res = dict(zip(SCN, ex.map(lambda kv: read_src(run(f'scn_{kv[0]}', kv[1])), SCN.items())))
    desc = {'A': 'API token 量年增取高端（' + ' '.join(TOK_HI) + '）', 'B': '基準', 'C': 'API token 量年增取低端（' + ' '.join(TOK_LO) + '）'}
    out = dict(note='scripts/val_snapshots.py scen 產生；值＝LibreOffice 重算後 Val 第 3 節 B 列（各情境改 Inputs 後）；勿手改',
               model=f'build_xlsx.py（{dt.date.today().isoformat()}）', scenarios={})
    for s in SCN:
        v = res[s]
        out['scenarios'][s] = dict(id=f'VAL_SCN_{s}', desc=desc[s], sets=SCN[s], values=v, check=sum(v['rev']) + sum(v['ebit']))
    json.dump(out, open(os.path.join(ROOT, 'data', 'val_scenarios.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    for s in SCN:
        print(s, {k: [round(x, 1) for x in v] for k, v in res[s].items()})


def sens():
    jobs = [('base', [])] + [(f'{c[0]}_lo', c[3]) for c in CASES] + [(f'{c[0]}_hi', c[4]) for c in CASES]
    with ThreadPoolExecutor(3) as ex:
        tp = dict(zip([j[0] for j in jobs], ex.map(lambda j: read_tp(run(f'sens_{j[0]}', j[1])), jobs)))
    out = dict(note='scripts/val_snapshots.py sens 產生；值＝改寫 Inputs 後整本 LibreOffice 重算的加權目標價（港元）；勿手改',
               model=f'build_xlsx.py（{dt.date.today().isoformat()}）', base_tp=tp['base'], cases=[])
    for k, name, plain, lo, hi, ll, hl, lp, hp in CASES:
        out['cases'].append(dict(key=k, name=name, plain=plain, tp_lo=tp[f'{k}_lo'], tp_hi=tp[f'{k}_hi'], lo_label=ll, hi_label=hl, lo_plain=lp, hi_plain=hp,
                                 lo_sets=lo, hi_sets=hi))
        print(f'{name}: {tp[k + "_lo"]:.2f} / {tp[k + "_hi"]:.2f}（基準 {tp["base"]:.2f}）')
    json.dump(out, open(os.path.join(ROOT, 'data', 'val_sens.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1)


if __name__ == '__main__':
    {'scen': scen, 'sens': sens}[sys.argv[1]]()
