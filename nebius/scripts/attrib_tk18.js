// v0.4a 步驟 7：v0.3 → v0.4 目標價變動拆解（已決定事項 12；HTML 引擎，與 Excel 經 cmp31 逐項一致）
// (a) 時間推移、(b) 實際數更新、(d) 方法變更皆 0；全部歸 (c) 假設變更中的「Tokenomics 同步」，再拆四段：
//   ① 換快照（v5.27 → v18.0；k 固定 v0.3 值：k_長約 0.76、k_現貨 1.76；每 MW 建置成本仍為 v0.3 寫死值 50.12／50.26）
//   → ② k 隨成本重算（k＝證據價格 ÷ 目前快照同世代成本）→ ③ 資本支出改引用 TK_CapexTotal（Σ 新增世代占比 × IF_CapexTotal）→ ④ 上限檢查改四層（只檢查，應為 0）
// 每一步期初可計費 MW 依「最新季營收 × 4 ÷ 首期每 MW 年收入」重算（與 Excel 活公式相同）。
// 用法：node scripts/attrib_tk18.js <v5.27 快照檔> [輸出.json]（v5.27 快照已自工作樹移除，自 git 歷史取出：git show <commit>:nebius/data/tokenomics_snapshot_v5.27.json > 暫存檔）
const fs = require('fs'), path = require('path');
const dir = path.join(__dirname, '..'); require(path.join(dir, 'load_engine.js'))(dir);
const toSnap = items => Object.fromEntries(Object.entries(items).filter(([, it]) => !it.missing).map(([n, it]) => [n, it.values]));
const NEW = COMPANY_DATA.tkSnap, OLD = toSnap(JSON.parse(fs.readFileSync(process.argv[2], 'utf8')).items);
for (const L of ['100', 'Util', 'Prod', 'Life']) OLD[`IFW_RevGWFleet_${L}`] = OLD.IF_RevGWFleet; // v5.27 無四層名稱：只影響上限檢查（不改數字），以 IF_RevGWFleet 代入使引擎可執行
const LEGACY_COST = COMPANY_DATA.scenarios.capexTemplate.costMW, K_FIX = { kLong: 0.76, kSpot: 1.76 }; // v0.3 的固定倍數與寫死建置成本
const run = (sc, st) => {
  COMPANY_DATA.tkSnap = st.snap;
  const rev = tkAnchorQ(sc, st.kfix ? K_FIX : {}).rev, s = scnQ(DEFAULTS, sc);
  s.m = { ...s.m, revMW: rev }; s.billableOpen = Math.round(BO_RUNRATE / rev[0]);
  s.a = { ...s.a, costMW: st.tkCapex ? capexTkQ(null) : [...LEGACY_COST] };
  const r = fA(s, null, VAL_DEFAULTS); COMPANY_DATA.tkSnap = NEW;
  return { tgt: r.tgt, eq: r.eq, gap: r.gap, rev0: rev[0] * 1e3, rev4: rev[4] * 1e3, bo: s.billableOpen, cost: s.a.costMW };
};
const STAGES = [{ snap: OLD, kfix: 1, tkCapex: 0 }, { snap: NEW, kfix: 1, tkCapex: 0 }, { snap: NEW, kfix: 0, tkCapex: 0 }, { snap: NEW, kfix: 0, tkCapex: 1 }];
const out = {};
for (const sc of ['low', 'base', 'high']) {
  const S = STAGES.map(st => run(sc, st)), d = i => S[i].tgt - S[i - 1].tgt;
  out[sc] = { steps: S, d1: d(1), d2: d(2), d3: d(3), d4: 0, total: S[3].tgt - S[0].tgt, v03: S[0].tgt, v04: S[3].tgt };
  console.log(`${sc}：v0.3 $${S[0].tgt.toFixed(2)} → ① 換快照 ${d(1) >= 0 ? '+' : ''}${d(1).toFixed(2)} → ② k 隨成本重算 ${d(2) >= 0 ? '+' : ''}${d(2).toFixed(2)} → ③ 資本支出改引用 TK ${d(3) >= 0 ? '+' : ''}${d(3).toFixed(2)} → ④ 上限檢查 0 → v0.4 $${S[3].tgt.toFixed(2)}；合計 ${out[sc].total >= 0 ? '+' : ''}${out[sc].total.toFixed(2)}（①＋②＋③＋④ ${(out[sc].d1 + out[sc].d2 + out[sc].d3).toFixed(2)}）`);
}
fs.writeFileSync(process.argv[3] || path.join(dir, 'out', 'attrib_tk18.json'), JSON.stringify(out, null, 1));
