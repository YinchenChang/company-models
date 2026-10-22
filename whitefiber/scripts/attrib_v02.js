// WhiteFiber v0.2 B 段步驟 7：v0.1 → v0.2 的變動拆解（已決定事項 12；HTML 引擎，與 Excel 經 cmp31 逐項一致）。
// 一、目標價逐步拆解（三情境各自加總＝總變動）：
//   ① 快照與錨取代舊推導（k＝1；保留 v0.2 的情境價格比）→ ② 套用定價倍數 k → ③ 情境軸分離（容量情境不再改價格）
//   → ④ 資本支出與汰換年限改引用 TK（Σ 新增世代占比 × IF_CapexIT）→ ⑤ EBITDAR 重校準（基準穩態 EBITDA 率維持 47%）→ ⑥ 上限檢查改四層（只檢查、不進估值，應為 0）
//   (a) 時間推移、(b) 實際數更新皆 0；①、④ 含 (c) TK 同步，②③⑤⑥ 為 (d) 方法變更（①為兩者合併：舊推導停用與新快照同時發生）。
// 二、每 MW 年收入（基準容量情境）逐年橋：舊值 17.40 → TK 錨本身變動 → 成本加成停用 → 市場價平均停用 → 世代組合 → k 證據 → 新值。
// 每一步期初可計費 MW 依「最新季經常性雲端營收 × 4 ÷ 首期每 MW 年收入」重算（與 Excel 活公式相同）。
// 用法：node scripts/attrib_v02.js [輸出.json]（v0.2 參考值取自 git 的 LEGACY_REF，預設 b4ba593）
const fs = require('fs'), path = require('path'), cp = require('child_process');
const dir = path.join(__dirname, '..'); require(path.join(dir, 'load_engine.js'))(dir);
const REF = process.env.LEGACY_REF || '1079d5c', K3 = ['low', 'base', 'high'], OLD = COMPANY_DATA.scenarios.revMW, ratio = sc => OLD[sc].map((x, i) => x / OLD.base[i]);
const ref = JSON.parse(cp.execFileSync('git', ['-C', dir, 'show', `${REF}:whitefiber/company.json`], { encoding: 'utf8' })), E_OLD = ref.defaults.ebitdarAdj, E_NEW = DEFAULTS.ebitdarAdj, COST_OLD = ref.scenarios.capexTemplate.costMW;
const LIFE_OLD = ref.defaults.gpuLife, OAR = COMPANY_DATA.scenarios.billableRatio.openAnnualRevenue, UF = UFQ / 1e3; // WhiteFiber：金額單位 US$m（UF＝1）；rev 為模型單位/MW
const run = (sc, st) => {
  const s = scnQ(DEFAULTS, sc), bo = RNDQ(OAR / st.rev[0]);
  s.m = { ...s.m, revMW: st.rev, billable: SCENARIOS[sc].acc.map((x, t) => RNDQ(bo + (x - bo) * SC_BR[t])) }; s.billableOpen = bo;
  s.a = { ...s.a, costMW: st.cost }; s.ebitdarAdj = st.eAdj; s.gpuLife = st.life ?? LIFE_OLD; // 汰換年限：v0.1 公司會計 5 年 → ④ 起引用 TK_DeprLifeIT（6 年）
  const r = fSensQ(s, VAL_DEFAULTS);
  return { tgt: r.tgt, eq: r.eq, gap: r.gap, levMax: r.levMax, rev0: st.rev[0], rev4: st.rev[4], bo, fy27Rev: r.fy27Rev };
};
// 舊方法（v0.2：路徑 A 成本加成與路徑 B 市場價 50/50 平均；資料 data/permw_tokenomics_20261008.json，GB300 基準成本情境）以新快照的 GB300 IF_HoldEcon 重算＝純 Tokenomics 同步（①a）；其餘（①b）＝以錨取代舊推導
const PMD = JSON.parse(fs.readFileSync(path.join(dir, 'data/permw_tokenomics_20261008.json'), 'utf8')), HE_NEW0 = COMPANY_DATA.tkSnap.IF_HoldEcon[`GB300 NVL72`][`基準`], GP0 = COMPANY_DATA.tkSnap.IF_GPUsPerGW[`GB300 NVL72`][`基準`] / 1e3,
  PA = { low: [0.0, 0.06], base: [0.15, 0.08], high: [0.30, 0.10] }, PB = { low: 9.7, base: 5.21 * GP0 * 8760 / 1e6 * .85, high: 7.82 * GP0 * 8760 / 1e6 * .90 }, // 路徑 A：[加成 k, 公司層級費用 o]；路徑 B：保守＝IREN 長約、基準＝spot 5.21 × 85%、積極＝24 個月預留 7.82 × 90%
  oldFormula = (sc, he) => (he * (1 + PA[sc][0]) / (1 - PA[sc][1]) + PB[sc]) / 2;
const out = { ref: REF, eOld: E_OLD, eNew: E_NEW, costOld: COST_OLD, costNew: capexTkQ(null), target: {}, permw: {} };
for (const sc of K3) {
  const R = ratio(sc), A1 = tkAnchorQ(sc, { kLong: 1, kSpot: 1 }).rev, T = tkAnchorQ(sc).rev, CN = capexTkQ(null);
  const S1a = Array(5).fill(oldFormula(sc, HE_NEW0)), S = [{ rev: OLD[sc], cost: COST_OLD, eAdj: E_OLD }, { rev: A1.map((x, i) => x * R[i]), cost: COST_OLD, eAdj: E_OLD }, { rev: T.map((x, i) => x * R[i]), cost: COST_OLD, eAdj: E_OLD },
    { rev: T, cost: COST_OLD, eAdj: E_OLD }, { rev: T, cost: CN, eAdj: E_OLD, life: DEFAULTS.gpuLife }, { rev: T, cost: CN, eAdj: E_NEW, life: DEFAULTS.gpuLife }].map(st => run(sc, st));
  const live = fSensQ(scnQ(DEFAULTS, sc), VAL_DEFAULTS).tgt, d = i => S[i].tgt - S[i - 1].tgt, R1a = run(sc, { rev: S1a, cost: COST_OLD, eAdj: E_OLD });
  out.target[sc] = { steps: S, step1a: R1a, d1a: R1a.tgt - S[0].tgt, d1b: S[1].tgt - R1a.tgt, rev1a: S1a[0], v02: S[0].tgt, d1: d(1), d2: d(2), d3: d(3), d4: d(4), d5: d(5), d6: 0, v03: S[5].tgt, live, total: S[5].tgt - S[0].tgt, sumSteps: d(1) + d(2) + d(3) + d(4) + d(5), err: S[5].tgt - live };
  const o = out.target[sc], f = x => (x >= 0 ? '+' : '') + x.toFixed(2);
  console.log(`${sc}：v0.2 $${o.v02.toFixed(2)} → ① ${f(o.d1)}〔①a TK 同步 ${f(o.d1a)}＋①b 錨取代舊推導 ${f(o.d1b)}〕→ ② ${f(o.d2)} → ③ ${f(o.d3)} → ④ ${f(o.d4)} → ⑤ ${f(o.d5)} → ⑥ 0 → v0.2 $${o.v03.toFixed(2)}（引擎現值 $${live.toFixed(2)}，差 ${(o.v03 - live).toExponential(1)}）；合計 ${f(o.total)}（①–⑤ 相加 ${f(o.sumSteps)}）；期初可計費 MW ${S.map(x => x.bo).join('／')}`);
}
// 二、每 MW 年收入橋（基準容量情境；v0.2 路徑 A 成本加成 ＋ 路徑 B 市場價 50/50 平均，GB300 基準成本情境；資料 data/permw_tokenomics_20261008.json＝v5.24 手抄值）
const PM = JSON.parse(fs.readFileSync(path.join(dir, 'data/permw_tokenomics_20261008.json'), 'utf8')), HE_OLD = 12.7247463559608, // v5.24 GB300 IF_HoldEcon 基準（model/archive/20261007_Tokenomics_v5.24.xlsx；permw 檔四捨五入為 12.72）
  HE_NEW = COMPANY_DATA.tkSnap.IF_HoldEcon[`GB300 NVL72`][`基準`], gpus = COMPANY_DATA.tkSnap.IF_GPUsPerGW[`GB300 NVL72`][`基準`] / 1e3, SPOT = 5.21, UTIL = .85,
  pathA = h => h * (1 + .15) / (1 - .08), pathB = SPOT * gpus * 8760 / 1e6 * UTIL, TB = tkAnchorQ(`base`);
out.permw.base = TB.rev.map((rev, t) => {
  const a = TB.anchor[t] * 1e3 / UFQ, k = TB.k[t], S0 = OLD.base[t] * 1e3 / UFQ, s0c = (pathA(HE_OLD) + pathB) / 2,
    s1 = (pathA(HE_NEW) + pathB) / 2, s2 = (HE_NEW + pathB) / 2, s3 = HE_NEW, s4 = a, s5 = a * k;
  return { period: PERIODS[t], old: S0, recompute: s0c, rounding: S0 - s0c, s1, s2, s3, s4, new: s5, d: { rounding: S0 - s0c, tkAnchor: s1 - s0c, costPlus: s2 - s1, marketPrice: s3 - s2, genMix: s4 - s3, kEvidence: s5 - s4 }, anchorAvg: a, k, sum: S0 + (s1 - s0c) + (s2 - s1) + (s3 - s2) + (s4 - s3) + (s5 - s4) }
});
out.permw.params = { HE_OLD, HE_NEW, gpus, SPOT, UTIL, pathA_old: pathA(HE_OLD), pathA_new: pathA(HE_NEW), pathB };
for (const x of out.permw.base) console.log(`每 MW ${x.period}：舊 ${x.old.toFixed(2)}｜進位 ${x.d.rounding.toFixed(3)}｜TK 錨 ${x.d.tkAnchor.toFixed(2)}｜成本加成停用 ${x.d.costPlus.toFixed(2)}｜市場價平均停用 ${x.d.marketPrice.toFixed(2)}｜世代組合 ${x.d.genMix.toFixed(2)}｜k 證據 ${x.d.kEvidence.toFixed(2)}｜新 ${x.new.toFixed(2)}｜加總（含進位）${x.sum.toFixed(3)}`);
fs.mkdirSync(path.join(dir, 'out'), { recursive: true });
fs.writeFileSync(process.argv[2] || path.join(dir, 'out', 'attrib_v02.json'), JSON.stringify(out, null, 1));
