# MAG v0.2.1（2026-10-22，自 Nebius v0.4／Oracle v0.3 的 kspec.py 改寫為 MAG 的 pricing.kLong／kSpot 結構）：
# 定價倍數 k 由「證據價格 ÷ 目前 Tokenomics 快照同世代成本（基準成本情境）」推算——錨改變時 k 跟著重算，價格（事實）不變。
# company.json → pricing.kLong／kSpot 的 low、base、high 可為：
#   數字（固定倍數）｜{"ref": 證據 id}（＝pricing.kEvidence 該筆價格 ÷ 同世代 tkName 基準值）｜{"priceEq": 價格, "gen": 世代代碼, "tkName"}（把判斷倍數換成固定價格後 ÷ 目前成本）
# 世代代碼：H100／GB200／GB300／VR200（與 Tokenomics 快照 source.generations 的 code 相同）。
# HTML 端同算式：segA.js kSpecApplyQ。用法：from kspec import resolve; resolve(co, 快照 dict) → 原地把規格換成數字（回傳 {kLong|kSpot: {low|base|high: 規格}}）。


def _tk(snap, n, g):
    code = {x['code']: x['name'] for x in snap['source']['generations']}
    return snap['items'][n]['values'][code[g]]['基準']


def resolve(co, snap):
    pr = co.get('pricing')
    if not pr or 'kEvidence' not in pr: return {}
    ev = {e['id']: e for e in pr['kEvidence']}
    def val(s):
        if isinstance(s, (int, float)): return s
        if 'ref' in s: e = ev[s['ref']]; return e['price'] / _tk(snap, e['tkName'], e['gen'])
        if 'priceEq' in s: return s['priceEq'] / _tk(snap, s['tkName'], s['gen'])
        raise SystemExit(f'pricing：無法辨識的 k 規格 {s}')
    spec = {}
    for side in ('kLong', 'kSpot'):
        for a in ('low', 'base', 'high'):
            spec.setdefault(side, {})[a] = pr[side][a]; pr[side][a] = val(pr[side][a])
    return spec
