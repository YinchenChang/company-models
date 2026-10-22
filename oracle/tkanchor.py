# Oracle v0.3（2026-10-22，自 Nebius v0.2a／v0.4 tkanchor.py 改寫）：每 MW 年收入以 Tokenomics 為錨（已決定事項 14、下游資料契約第 4 條）——建置時的參考計算。
# 與 Excel「每MW收入_錨定」分頁（活公式）及 HTML 引擎 segA.js → tkAnchorQ() 同一算式；本檔用於：
#   (1) build_xlsx.py 檢查 company.json → defaults.m.revMW／defaults.billableOpen／defaults.m.billable 等於預設情境的錨定值；
#   (2) scripts/tkanchor_sync.py --write 回寫上述預設值；(3) 報告與對照 Excel。
# 算式（各容量情境 sc、價格情境 px）：
#   期末在役 MW_t＝已連網 MW_t（scenarios.mwPath；權重用已連網、不用可計費，避免與期初可計費 MW 校準循環）；期初（首期）＝fleet.openMw；新增＝MAX(0, 期末 − 期初)
#   各世代期末 MW＝期初世代 MW＋新增 × newMix_t；平均在役世代占比＝（期初_g＋期末_g）÷（期初＋期末）
#   錨_t＝Σ 占比 × IF_HoldEcon_g（成本情境，預設基準）÷ 1000（US$bn/MW·年）
#   k_t＝隨需占比 × k_現貨 ＋（1 − 隨需占比）× k_長約；每 MW 年收入_t＝錨_t × k_t（100% 計費時數；利用率不另扣）
import json, os

HERE = os.path.dirname(os.path.abspath(__file__))


def load(co=None):
    co = co or json.load(open(os.path.join(HERE, 'company.json'), encoding='utf-8'))
    tk = json.load(open(os.path.join(HERE, co['tokenomics']['snapshotFile']), encoding='utf-8'))['items']
    import sys; sys.path.insert(0, HERE); import kspec; kspec.resolve(co, tk)  # k 規格（證據價格 ÷ 目前快照）→ 數字
    return co, tk


def acc_path(co, sc, years):
    P = co['scenarios']['mwPath']; a = []
    for i, L in enumerate(years):
        a.append(min(P['contracted'][sc][i], P['connectedStart'] if i == 0 else a[i - 1] + P['pace'][sc] * L))
    return a


def compute(co, tk, sc, years, px='base', tkCase='基準', kLong=None, kSpot=None, newMix=None, od=None):
    AM = co['pricing']['anchorMultiple']; FL = co['fleet']; G = FL['generations']
    end = acc_path(co, sc, years)
    nm = newMix or FL['newMix']
    prev = {g: FL['openMw'] * FL['openMix']['mix'].get(g, 0) for g in G}; prevT = FL['openMw']
    ref = lambda side: next(x for x in AM['evidence'] if x['label'] == AM[side]['refEvidence'])
    rr = lambda x: 1 if tkCase == '基準' else tk[x['tkName']]['values'][x['gen']]['基準'] / tk[x['tkName']]['values'][x['gen']][tkCase]  # 非基準成本情境：k 以證據世代成本比例重算（價格是事實）
    kL = (AM['long'][px] if kLong is None else kLong) * rr(ref('long'))
    kS = (AM['spot'][px] if kSpot is None else kSpot) * rr(ref('spot'))
    odv = AM['onDemandShare']['base'] if od is None else od
    he = {g: tk['IF_HoldEcon']['values'][g][tkCase] for g in G}
    rf = {g: tk['IF_RevGWFleet']['values'][g][tkCase] for g in G}
    LAYERS = ['100', 'Util', 'Prod', 'Life']; LAY = (co.get('methodology', {}).get('checks') or {}).get('revCapLayer', 'Util')
    lv = {L: {g: tk[f'IFW_RevGWFleet_{L}']['values'][g][tkCase] for g in G} for L in LAYERS}
    R = {k: [] for k in ('start', 'end', 'add', 'avg', 'anchor', 'k', 'rev', 'capRef', 'capRatio')}
    R['share'] = {g: [] for g in G}; R['kL'] = kL; R['kS'] = kS; R['capLay'] = {L: [] for L in LAYERS}; R['capRL'] = {L: [] for L in LAYERS}
    for t in range(5):
        add = max(0.0, end[t] - prevT)
        cur = {g: prev[g] + add * nm[t].get(g, 0) for g in G}
        tot = prevT + end[t]
        sh = {g: (prev[g] + cur[g]) / tot for g in G}
        A = sum(sh[g] * he[g] for g in G) / 1000
        k = odv * kS + (1 - odv) * kL
        cap = sum(sh[g] * rf[g] for g in G) / 1000
        lay = {L: sum(sh[g] * lv[L][g] for g in G) / 1000 for L in LAYERS}
        for L in LAYERS: R['capLay'][L].append(lay[L]); R['capRL'][L].append(A * k / lay[L])
        for key, v in (('start', prevT), ('end', end[t]), ('add', add), ('avg', tot / 2), ('anchor', A), ('k', k),
                       ('rev', A * k), ('capRef', cap), ('capRatio', A * k / lay[LAY])):
            R[key].append(v)
        for g in G: R['share'][g].append(sh[g])
        prev, prevT = cur, end[t]
    return R


def bopen(co, rev0):
    return int(co['scenarios']['billableRatio']['openAnnualRevenue'] / rev0 + 0.5)


def billable_path(co, sc, years, bo):
    ac = acc_path(co, sc, years); br = co['scenarios']['billableRatio']['ratio']
    return [int(bo + (x - bo) * b + 0.5) for x, b in zip(ac, br)]


if __name__ == '__main__':
    import calendar_q
    co = calendar_q.load(HERE); _, tk = load(co)
    for sc in ('low', 'base', 'high'):
        R = compute(co, tk, sc, co['periodYears'])
        print(sc, 'anchor', [round(x * 1e3, 3) for x in R['anchor']], 'k', [round(x, 3) for x in R['k']], 'rev', [round(x * 1e3, 3) for x in R['rev']],
              'bo', bopen(co, R['rev'][0]), 'cap', [round(x, 2) for x in R['capRatio']])
