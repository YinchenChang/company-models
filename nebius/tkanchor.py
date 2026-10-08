# v0.2a（2026-10-08）：每 MW 年收入以 Tokenomics 為錨（已決定事項 14、下游資料契約第 4 條）——建置時的參考計算。
# 與 Excel「每MW收入_錨定」分頁（活公式）及 HTML 引擎 segA.js → tkAnchorQ() 同一算式；本檔用於：
#   (1) build_xlsx.py 檢查 company.json → defaults.m.revMW／defaults.billableOpen 等於預設情境的錨定值；
#   (2) scripts/tkanchor_sync.py --write 回寫上述預設值；(3) 報告與對照 Excel。
# 算式（各容量情境 sc、價格情境 px）：
#   期末在役 MW_t＝已連網 MW_t × 在役比例_t；期初（首期）＝priceCheck.inServiceMw；新增＝MAX(0, 期末 − 期初)
#   各世代期末 MW＝期初世代 MW＋新增 × newMix_t；平均在役世代占比＝（期初_g＋期末_g）÷（期初＋期末）
#   錨_t＝Σ 占比 × IF_HoldEcon_g（成本情境，預設基準）÷ 1000（US$bn/MW·年）
#   長約 MW_t＝Σ 起始期 ≤ t 的合約 MW；長約占比_t＝MIN(1, 長約 MW_t ÷ 平均在役 MW_t)
#   k_t＝長約占比 × k_長約 ＋（1 − 長約占比）× k_現貨；每 MW 年收入_t＝錨_t × k_t（100% 計費時數；利用率不另扣）
import json, os

HERE = os.path.dirname(os.path.abspath(__file__))
PX = {'low': 'low', 'base': 'base', 'high': 'high'}


def load(co=None):
    co = co or json.load(open(os.path.join(HERE, 'company.json'), encoding='utf-8'))
    tk = json.load(open(os.path.join(HERE, co['tokenomics']['snapshotFile']), encoding='utf-8'))['items']
    return co, tk


def acc_path(co, sc, years):
    P = co['scenarios']['mwPath']; a = []
    for i, L in enumerate(years):
        a.append(min(P['contracted'][sc][i], P['connectedStart'] if i == 0 else a[i - 1] + P['pace'][sc] * L))
    return a


def compute(co, tk, sc, years, px='base', tkCase='基準', ls='mw', kLong=None, kSpot=None, newMix=None):
    AM = co['pricing']['anchorMultiple']; FL = co['fleet']; G = FL['generations']
    acc = acc_path(co, sc, years); br = co['scenarios']['billableRatio']['ratio']
    end = [acc[i] * br[i] for i in range(5)]
    st0 = co['priceCheck']['inServiceMw']
    nm = newMix or FL['newMix']
    prev = {g: st0 * FL['openMix']['mix'].get(g, 0) for g in G}; prevT = st0
    kL = AM['long'][px] if kLong is None else kLong
    kS = AM['spot'][px] if kSpot is None else kSpot
    he = {g: tk['IF_HoldEcon']['values'][g][tkCase] for g in G}
    rf = {g: tk['IF_RevGWFleet']['values'][g][tkCase] for g in G}
    R = {k: [] for k in ('start', 'end', 'add', 'avg', 'anchor', 'longMw', 'ls', 'k', 'rev', 'capRef', 'capRatio')}
    R['share'] = {g: [] for g in G}; R['kL'] = kL; R['kS'] = kS
    for t in range(5):
        add = max(0.0, end[t] - prevT)
        cur = {g: prev[g] + add * nm[t].get(g, 0) for g in G}
        tot = prevT + end[t]
        sh = {g: (prev[g] + cur[g]) / tot for g in G}
        A = sum(sh[g] * he[g] for g in G) / 1000
        avg = tot / 2
        lm = sum(c[ls] for c in AM['longShare']['contracts'] if c['start'] <= t)
        s = min(1.0, lm / avg)
        k = s * kL + (1 - s) * kS
        cap = sum(sh[g] * rf[g] for g in G) / 1000
        for key, v in (('start', prevT), ('end', end[t]), ('add', add), ('avg', avg), ('anchor', A), ('longMw', lm), ('ls', s), ('k', k),
                       ('rev', A * k), ('capRef', cap), ('capRatio', A * k / cap)):
            R[key].append(v)
        for g in G: R['share'][g].append(sh[g])
        prev, prevT = cur, end[t]
    return R


def billable_open(co, rev0):
    return int(co['latestQuarter']['revenue'] * 4 / rev0 + 0.5)


if __name__ == '__main__':
    import calendar_q
    co = calendar_q.load(HERE); _, tk = load(co)
    for sc in ('low', 'base', 'high'):
        R = compute(co, tk, sc, co['periodYears'])
        print(sc, 'anchor', [round(x * 1e3, 3) for x in R['anchor']], 'ls', [round(x, 3) for x in R['ls']],
              'k', [round(x, 3) for x in R['k']], 'rev', [round(x * 1e3, 3) for x in R['rev']], 'bo', billable_open(co, R['rev'][0]),
              'cap', [round(x, 2) for x in R['capRatio']])
