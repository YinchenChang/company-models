// v0.2a 步驟 7：v0.2 → v0.3 目標價變動 (d) 方法變更再拆（已決定事項 12；HTML 引擎，與 Excel 經 cmp31 逐項一致）
// ① 錨取代舊推導（k＝1；保留 v0.2 的情境價格比）→ ② 套用定價倍數 k → ③ 情境軸分離（每 MW 收入不再隨容量情境同向；積極不再用高價）→ ④ 上限檢查（只警示，應為 0）
// 每一步期初可計費 MW 依「最新季營收 × 4 ÷ 首期每 MW 年收入」重算（與 Excel 活公式相同）。
// 用法：node scripts/attrib_tkanchor.js [輸出.json]
const fs = require('fs'), path = require('path');
const dir = path.join(__dirname, '..'); require(path.join(dir, 'load_engine.js'))(dir);
const K3 = ['low', 'base', 'high'], OLD = COMPANY_DATA.scenarios.revMW, ratio = sc => OLD[sc].map((x, i) => x / OLD.base[i]);
const run = (sc, rev) => { const s = scnQ(DEFAULTS, sc); s.m = { ...s.m, revMW: rev }; s.billableOpen = Math.round(BO_RUNRATE / rev[0]); const r = fA(s, null, VAL_DEFAULTS); return { tgt: r.tgt, eq: r.eq, gap: r.gap, rev30: rev[4] * 1e3, bo: s.billableOpen }; };
const out = {};
for (const sc of K3) {
  const R = ratio(sc), A1 = tkAnchorQ(sc, { kLong: 1, kSpot: 1 }).rev, A2 = tkAnchorQ(sc).rev;
  const S = [run(sc, OLD[sc]), run(sc, A1.map((x, i) => x * R[i])), run(sc, A2.map((x, i) => x * R[i])), run(sc, A2)];
  out[sc] = { steps: S, d1: S[1].tgt - S[0].tgt, d2: S[2].tgt - S[1].tgt, d3: S[3].tgt - S[2].tgt, d4: 0, total: S[3].tgt - S[0].tgt };
  console.log(`${sc}：v0.2 $${S[0].tgt.toFixed(2)} → ① $${S[1].tgt.toFixed(2)}（${out[sc].d1.toFixed(2)}）→ ② $${S[2].tgt.toFixed(2)}（${out[sc].d2.toFixed(2)}）→ ③ $${S[3].tgt.toFixed(2)}（${out[sc].d3.toFixed(2)}）→ ④ 0；FY30 每 MW ${S.map(x => x.rev30.toFixed(2)).join('／')}；期初可計費 ${S.map(x => x.bo).join('／')}`);
}
fs.writeFileSync(process.argv[2] || path.join(dir, 'out', 'attrib_tkanchor.json'), JSON.stringify(out, null, 1));
