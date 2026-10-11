# W6 r1（2026-10-09）：定價倍數 k 的基準與區間由證據價格 ÷ 目前 Tokenomics 快照（基準成本情境）推算——錨改變時 k 跟著重算，價格（事實）不變。
# company.json → pricing.anchorMultiple.long／spot 的 base、low、high、sensMedian 可為：
#   數字（固定倍數）｜{"ref": 證據 label}（＝該筆價格 ÷ 同世代 tkName 基準值）｜{"priceEq": 價格, "gen", "tkName"}（區間端點換成固定價格）｜{"median": [證據 label…]}（各筆倍數的中位數）
# HTML 端同算式：segA.js kSpecQ。用法：from kspec import resolve; resolve(co, snapshot_items) → 原地把規格換成數字（回傳 {side: {key: 規格說明}}）。
import statistics


def _tk(items, n, g):
    return items[n]['values'][g]['基準']


def resolve(co, items):
    am = (co.get('pricing') or {}).get('anchorMultiple')
    if not am: return {}
    ev = {e['label']: e for e in am['evidence']}
    mult = lambda e: e['price'] / _tk(items, e['tkName'], e['gen'])
    def val(s):
        if isinstance(s, (int, float)): return s
        if 'ref' in s: return mult(ev[s['ref']])
        if 'priceEq' in s: return s['priceEq'] / _tk(items, s['tkName'], s['gen'])
        if 'median' in s: return statistics.median(mult(ev[x]) for x in s['median'])
        raise SystemExit(f'pricing.anchorMultiple：無法辨識的 k 規格 {s}')
    spec = {}
    for side in ('long', 'spot'):
        for k in ('base', 'low', 'high', 'sensMedian'):
            if k in am[side]:
                spec.setdefault(side, {})[k] = am[side][k]; am[side][k] = val(am[side][k])
    return spec
