const fs=require('fs'); require('./load_engine.js')(); // v4.0：引擎資料來自 company.json
// v4.5：Excel 列名稱中的期間字樣由日曆產生（佔位符 → 目前日曆字樣，與 build_xlsx 相同的 calendar_q.tokens）
const T=s=>Object.keys(CALQ.tokens).sort((a,b)=>b.length-a.length).reduce((x,k)=>x.split(k).join(CALQ.tokens[k]),s);
const SC=process.argv[2], AK=+(process.argv[3]||3), AKX=process.argv[3]!==undefined; // AK：EV/EBITDA 錨定年度（1＝FY27…4＝FY30）
const X=JSON.parse(fs.readFileSync('xl17_'+({low:1,base:2,high:3}[SC])+(AKX?'_a'+AK:'')+'.json','utf8'));
VAL_DEFAULTS.evYear=AK;
const q=structuredClone(DEFAULTS); q.scenario=SC; q.a=structuredClone(SCENARIOS[SC].a); q.mw31=SCENARIOS[SC].mw31; q.cvCap=SCENARIOS[SC].cvCap; q.m.accepted=[...SCENARIOS[SC].acc]; q.m.billable=[...SCENARIOS[SC].bil]; q.m.revMW=[...SCENARIOS[SC].rev];
const d=runFunding(q), p=runValuation(d,q,VAL_DEFAULTS), y=d.years, f=p.fwd;
const H=(k)=>y.map(e=>e[k]);
const rows=[];
function cmp(label, xlKey, html, note){
  const MAP={'產能與收入|':['運營_產能與收入|'],'支出與資金|':['各期收支|'],'損益與評價|':['損益|','評價_DCF與目標價|'],'輸入|':['輸入與假設|'],'債務明細|':['資產負債_既有債務|'],'站點租賃|':['運營_站點|'],'可比公司|':['評價_可比公司|'],'連動檢查|':['檢查_連動|']};
  let x=X[xlKey]; if(!x){for(const [o,ns] of Object.entries(MAP)){if(xlKey.startsWith(o)){for(const n of ns){const k=n+xlKey.slice(o.length); if(X[k]){x=X[k];break;}}}}}
  if(!x){rows.push([label,'MISSING XL KEY '+xlKey]);return;}
  const xv=x.slice(0,html.length).map(v=>typeof v==='number'?v:NaN);
  const diff=html.map((h,i)=>Math.abs((h??NaN)-(xv[i]??NaN)));
  const ok=diff.every(d=>d<0.005||(!Number.isFinite(d)&&false));
  rows.push([ok?'OK ':'XX ',label, html.map(v=>+(+v).toFixed(3)).join('/'), xv.map(v=>+(+v).toFixed(3)).join('/'), note||'']);
}
const S='產能與收入|', F='支出與資金|', V='損益與評價|';
cmp('Accepted', S+'Accepted MW（已驗收，單調不減）', d.m.accepted);
cmp('Billable', S+'Billable MW（上限為 Accepted）', d.m.billable);
cmp('平均在役MW', S+T('平均在役 MW（«P0» 欄為«STUBW»平均）'), H('avgBillable'));
cmp('容量上限', S+'容量上限（模型期最多能交付的收入）', H('capacity'));
cmp('排程RPO', S+'排程 RPO（模型期，依季報桶分攤）', H('scheduled'));
cmp('瓶頸', S+'產能瓶頸（模型期：排程 − 容量）', H('bottleneck'));
cmp('RPO轉換', S+'期初 RPO 轉換收入（模型期）', H('revenue'));
cmp('新簽約收入', S+'新簽約收入', H('newRev'));
cmp('未售', S+'未售產能（浪費）', H('unsold'));
cmp('isRev', S+'損益用算力收入（模型期＝RPO 轉換＋新簽約）', H('isRev'));
cmp('營收−MW×單價×利用率', S+'核對：算力收入 − 平均在役 MW × 每 MW × 利用率 × 期間', y.map(e=>e.isRev-e.capacity)); // v0.1b：MW 驅動時為 0
cmp('每MW年收入', '輸入|每 MW 年收入', d.m.revMW);
cmp('信用損失', S+'信用損失（期初 RPO 部分）', H('loss'));
cmp('RPO現金', S+'RPO 現金（可支應資本用途）', H('rpoCash'));
cmp('新簽約現金', S+'新簽約現金', H('newCash'));
cmp('EBITDA率', S+'EBITDA 率（損益與資金共用）', H('ebM'));
cmp('EBITDAR率', S+'EBITDAR 率（EBITDA 率＋租金÷營收）', H('cashMargin'));
cmp('模型期總營收', S+'模型期總營收（算力＋服務）', H('totRev'));
cmp('服務現金', '輸入|非算力服務現金', H('legacy'));
cmp('FY26全年營收(產能頁)', S+T('«P0» 全年總營收（«YTD» 實際＋模型期算力＋服務）'), [y[0].fyRevenue+y[0].servicesRev]);
cmp('毛CapEx模型期', '輸入|毛 CapEx（模型期，下游引用此列）', H('gross'));
cmp('全年CapEx公式', '輸入|全年毛 CapEx（公式）', H('capexFull'));
cmp('成長型', '輸入|成長型 CapEx（模型期）', H('capexGrowth'));
cmp('汰換', '輸入|GPU 汰換 CapEx', H('refresh'));
cmp('D&A車隊', '輸入|D&A（車隊折舊）', H('daFleet'));
cmp('表外租金', '輸入|表外現金租金（未起租）', q.a.newLease);
cmp('存量利息', '輸入|存量債務利息（下游引用此列）', H('intStock'));
cmp('JV合計(模型期)', '輸入|JV／策略投資出資', H('div'));
cmp('新債利息', F+'③ 新債利息（瀑布）', H('newDebtInt'));
cmp('融資前現金', F+'融資前現金（扣既有新債利息）', H('preCash'));
cmp('融資需求', F+'融資需求（補足至最低現金）', H('need'));
cmp('期末backlog', F+'期末 backlog', H('backlogEnd'));
cmp('新債舉借', F+'新債舉借', H('newDebt'));
cmp('可轉債發行', F+'可轉債發行', H('convNew')); cmp('可轉債期末', F+'可轉債（瀑布）期末餘額', H('convEnd')); // v0.1b
cmp('股權募資', F+'股權募資', H('equity'));
cmp('高息債', F+'高息債舉借（溢出）', H('junk'));
cmp('融資前累積現金', F+'融資前累積現金', H('preFinCum'));
cmp('累計新股', F+'累計新股', H('cumNewShares'));
cmp('總債務', F+'期末總債務（既有＋可轉債＋新債＋高息債）', H('totalDebtEnd'));
cmp('租金合計', F+'　租金合計', H('lease'));
cmp('排程還本(FY26含1H)', F+'⑤ 排程還本（季報到期表）', y.map((e,i)=>i===0?e.fyDebtPay:e.debtPay));
cmp('客戶預付', F+T('Ⓓ 客戶預付（«STUB» 起）'), H('external'));
cmp('預付認列', F+'　預付認列（非現金營收）', H('prepayRecog')); cmp('合約負債期末', F+'　合約負債期末', H('clEnd')); cmp('合約負債期初', F+'　合約負債期初（客戶預付餘額）', H('clBeg')); // v0.1b
cmp('營運來源合計', F+'營運來源合計', y.map((e,i)=>i===0?e.fySourcesOp:e.sourcesOp));
cmp('營運缺口', F+'營運缺口（不含融資、不含還本）', y.map((e,i)=>i===0?e.fyOperatingGap:e.operatingGap));
cmp('期末累積現金', F+'期末累積現金', H('cum'));
const HTMLuses=y.map((e,i)=>i===0? e.fyUsesCash : e.uses);
cmp('用途合計(HTML各期結果口徑)', F+'用途合計（現金口徑）', HTMLuses, 'HTML FY26 用認列 CapEx 且再加 1H 利息/租金');
cmp('JV(FY26含1H)', F+'④ JV／策略投資出資', y.map((e,i)=>i===0?e.fyDiv:e.div));
cmp('股權/可轉債', F+'Ⓔ 股權／可轉債（融資）', y.map((e,i)=>i===0?e.fyEquity:e.atm));
cmp('總來源', F+'總來源（含融資）', y.map((e,i)=>i===0?e.fySrcTotal:e.sources));
cmp('CFO(1H)', F+T('Ⓐ0 «YTDL» 實際營運現金流（CFO）'), y.map((e,i)=>i===0?e.fyCfo:0));
cmp('CapEx用途', F+T('① CapEx（用途用：«YTD» 現金／«STUB» 毛額）'), y.map((e,i)=>i===0?e.fyCapexUse:e.gross));
cmp('調節', F+T('　«YTD» 其他／受限現金調節'), y.map((e,i)=>i===0?e.hPlug:0));
cmp('期初累積現金', F+'　期初累積現金', y.map((e,i)=>i===0?ACTUAL_1H.cash1231:y[i-1].cum));
cmp('全年總營收', V+'全年總營收', f.map(e=>e.fyRevenue));
cmp('全年EBIT', V+'全年 GAAP 營業利益', f.map(e=>e.fyOpInc));
cmp('全年D&A', V+'全年 D&A', f.map(e=>e.fyDa));
cmp('全年EBITDA', V+'全年 EBITDA', f.map(e=>e.fyEbitda));
cmp('全年利息', V+'全年利息（含瀑布新債）', y.map((e,i)=>i===0?ACTUAL_1H.interest+e.interest:e.interest));
cmp('全年淨利', V+'全年淨利', f.map(e=>e.fyNi));
cmp('全年EPS', V+'全年 GAAP EPS', f.map(e=>e.fyEps));
cmp('全年EPS加回SBC', V+'全年 EPS（加回 SBC）', f.map(e=>e.fyNgEps));
// P&L
cmp('算力收入(損益)', V+'算力收入', f.map(e=>e.gpu));
cmp('總營收(損益,模型期)', V+'總營收', f.map(e=>e.revenue));
cmp('EBIT(模型期)', V+'營業利益（EBIT）', f.map(e=>e.opInc));
cmp('DCF股權', V+'股權價值', [p.d.equity]);
cmp('稅', V+'所得稅', f.map(e=>e.tax));
cmp('淨利(模型期)', V+'淨利', f.map(e=>e.ni));
cmp('D&A', V+'D&A', f.map(e=>e.da));
cmp('EBITDA', V+'EBITDA', f.map(e=>e.ebitda));
cmp('EPS(模型期)', V+'每股盈餘（EPS，模型期）', f.map(e=>e.eps));
cmp('股數', V+'股數（含瀑布新股）', f.map(e=>e.shares));
cmp('UFCF', V+'UFCF', f.map(e=>e.ufcf));
cmp('UFCF現值', V+'UFCF 現值', p.d.pv);
cmp('DCF每股', V+'DCF 每股', [p.d.invalid?0:p.d.perShare]); // v0.1b：DCF 失效時 HTML 為 NaN、Excel 為 0（與反向 DCF 同一口徑）
const NB='資產負債_新債與新股|';
cmp('BS 期末現金', NB+'期末現金', H('cum'));
cmp('BS 營運收支淨額', NB+'營運收支淨額（含期後股權／可轉債，不含瀑布）', H('preFinGap'));
cmp('BS 總債務', NB+'總債務', H('totalDebtEnd'));
cmp('BS 淨負債', NB+'淨負債（總債務 − 期末現金）', y.map(e=>e.totalDebtEnd-e.cum));
cmp('BS 總股數', NB+'總股數（期末）', f.map(e=>e.shares));
cmp('BS 債務/EBITDA', NB+'總債務 ÷ EBITDA（年化）', y.map((e,t)=>e.totalDebtEnd/Math.max(e.ebitdaPL/PERIOD_YEARS[t],.01)));
cmp('DCF 0截斷', V+'DCF 每股：0 截斷', [p.d.zeroPerShare]);
cmp('DCF 選擇權', V+'DCF 每股：選擇權（Merton）', [p.d.optPerShare]);
cmp('DCF 失效', V+'DCF 失效？（WACC ≤ g 或常態化 FCF ≤ 0）', [p.d.invalid?1:0]);
cmp('錨定年EBITDA', V+'錨定年度 EBITDA', [f[p.evK].ebitda]);
cmp('錨定年末淨負債', V+'錨定年度末淨負債（總債務 − 現金）', [p.ndA]);
cmp('錨定年末股數', V+'錨定年度末股數（含瀑布新股）', [p.shA]);
// v0.1b：可轉債稀釋（分類、股數、淨負債、還本、票息、期末餘額）
const DB='資產負債_既有債務|';
cmp('可轉債還本', DB+'可轉債到期還本（債務處理）', H('cvAmort')); cmp('可轉債期末餘額', DB+'可轉債期末餘額（債務處理）', H('cvEnd')); cmp('可轉債票息', DB+'可轉債票息（債務處理）', H('cvInt'));
cmp('可轉債第一輪目標價', V+'第一輪：加權目標價', [p.cv.tp1]); cmp('第二輪判斷價', V+'第二輪判斷價', [p.cv.px2]);
cmp('第一輪轉股股數', V+'第一輪：轉股股數', [p.cv.sh1-p.shares0]); cmp('第二輪轉股股數', V+'第二輪：轉股股數', [p.cv.sh2-p.shares0]);
cmp('評價股數', V+'評價股數（含可轉債轉股）', [p.shares]); cmp('評價淨負債', V+'評價淨負債（含可轉債與調整項）', [p.v.netDebt]);
cmp('債務處理到期本金', V+'第二輪：債務處理到期本金', [p.cv.m2]);
cmp('其他事業EBITDA', F+T('Ⓒ2 其他事業 EBITDA（«STUB» 起）'), H('otherEbitda')); cmp('淨負債調整項', '輸入與假設|淨負債調整項（類債 − 持股 ×（1 − 折價））', [ndAdjQ(VAL_DEFAULTS)]); // v0.1b 步驟 6
cmp('PV股權', V+'新股募得現金（現值）', [p.d.pvEquityRaised]);
cmp('融資後股數', V+'融資後股數（含瀑布新股）', [p.d.postShares]);
cmp('EV/EBITDA融資後', V+'EV/EBITDA 每股（融資後）', [p.peAdj]);
cmp('加權目標價', V+'加權目標價', [p.call.blended]);
// v4.1：目標價區間（情境區間、方法區間、100% EV/EBITDA 目標價）與判斷句、DCF 權重說明（文字逐字比對）
{ const R=targetRange(d,q,VAL_DEFAULTS,p), RG='評價_DCF與目標價|目標價區間｜', CODE={買進:1,中立:0,賣出:-1};
  cmp('區間 點位', RG+'點位（目前輸入的加權目標價）', [R.pt]);
  cmp('區間 點位評等', RG+'點位評等代碼（1＝買進、0＝中立、−1＝賣出）', [CODE[p.call.call]]);
  cmp('區間 賣出門檻', RG+'賣出門檻價（現價 × (1 − '+pctQ(SELL_TH)+')）', [R.th]);
  R.scen.forEach(s=>cmp('區間 情境 '+s.sc, RG+SCENARIOS[s.sc].label+'（加權目標價／評等代碼／距賣出門檻）', [s.tgt, CODE[s.call], s.gap]));
  cmp('情境區間 A', RG+'情境區間（下緣／上緣）', R.A);
  cmp('方法區間倍數', RG+'方法區間倍數（下端／上端）', R.mults);
  cmp('方法區間 B', RG+'方法區間（下緣／上緣）', R.B);
  cmp('100% EV/EBITDA 目標價', RG+'100% EV/EBITDA 目標價（價位／評等代碼）', [R.ev100, CODE[R.ev100Call]]);
  cmp('未截斷 DCF 每股', RG+'未截斷 DCF 每股', [p.d.perShareRaw]);
  const cmpT=(label,xlKey,h)=>{ const x=X[xlKey]; if(!x){rows.push([label,'MISSING XL KEY '+xlKey]);return;} rows.push([x[0]===h?'OK ':'XX ',label,JSON.stringify(h),JSON.stringify(x[0])]); };
  cmpT('文字 點位摘要', RG+'點位摘要', R.head);
  cmpT('文字 方法區間標籤', RG+'方法區間標籤', R.bLabel);
  cmpT('文字 點位位置', RG+'點位位置', R.pos);
  cmpT('文字 判斷句', RG+'判斷句', R.judge);
  cmpT('文字 DCF 權重說明', '評價_DCF與目標價|DCF 權重說明', R.note); }
// v4.3：市場共識（I 區逐筆：數值、文字、擷取日期、標記）與一頁摘要（差距、隱含倍數、結論句／判斷句／隱含倍數句／Q3 句逐字）
{ const R=targetRange(d,q,VAL_DEFAULTS,p), cv=consensusView(d,p,VAL_DEFAULTS,R,q), SM='摘要|';
  const cmpT=(label,xlKey,h)=>{ const x=X[xlKey]; if(!x){rows.push([label,'MISSING XL KEY '+xlKey]);return;} rows.push([x[0]===h?'OK ':'XX ',label,JSON.stringify(h),JSON.stringify(x[0])]); };
  for(const it of consensusItems()){ const k='輸入與假設|'+it.label, x=X[k]; if(!x){rows.push(['共識 '+it.label,'MISSING XL KEY '+k]);continue;}
    const okV=it.vals.length?it.vals.every((z,i)=>typeof x[i]==='number'&&Math.abs(z-x[i])<1e-9):x[0]===it.text;
    const ok=okV&&x[3]===it.date&&(x[4]??'')===it.tag;
    rows.push([ok?'OK ':'XX ','共識 '+it.label,JSON.stringify([...it.vals,it.text,it.date,it.tag]),JSON.stringify(x)]); }
  for(const [n,k] of [['營收','rev'],['調整後 EBITDA','ebitda'],['CapEx（毛額）','capex'],['淨負債','nd']]){
    cmp('摘要 '+n+' 模型',SM+'差異｜'+n+'｜模型',cv.rows.map(r=>r.m[k])); cmp('摘要 '+n+' 共識',SM+'差異｜'+n+'｜共識',cv.rows.map(r=>r.c[k]));
    cmp('摘要 '+n+' 差距',SM+'差異｜'+n+'｜差距',cv.rows.map(r=>r.gap[k]));
    if(k==='ebitda'||k==='nd') cmp('摘要 '+n+' 差距金額',SM+'差異｜'+n+'｜差距（金額）',cv.rows.map(r=>r.dif[k])); }
  cmp('摘要 EBITDA率 模型',SM+'差異｜EBITDA 率｜模型',cv.rows.map(r=>r.m.ebM)); cmp('摘要 EBITDA率 共識',SM+'差異｜EBITDA 率｜共識',cv.rows.map(r=>r.c.ebM));
  cmp('摘要 EBITDA率 差距pt',SM+'差異｜EBITDA 率｜差距（百分點）',cv.rows.map(r=>r.gap.ebM));
  cmp('摘要 分歧起始年',SM+'差異｜分歧起始年（1＝FY26…3＝FY28；0＝無）',[cv.first+1]);
  cmp('隱含 共識目標價倍數',SM+'隱含｜共識平均目標價隱含 FY28 EV/EBITDA',[cv.impTgt]); cmp('隱含 現價倍數',SM+'隱含｜現價隱含 FY28 EV/EBITDA',[cv.impPx]);
  cmp('隱含 模型上緣',SM+'隱含｜模型方法區間上緣',[cv.mHi]);
  cmp('摘要 點位',SM+'結論｜點位（加權目標價）',[R.pt]); cmp('摘要 空間',SM+'結論｜空間',[cv.up]); cmp('摘要 點位−門檻',SM+'結論｜點位 − 賣出門檻',[cv.gapTh]);
  cmpT('文字 摘要評等',SM+'結論｜評等',p.call.call); cmpT('文字 摘要結論句',SM+'結論｜結論句',cv.head); cmpT('文字 摘要情境判斷句',SM+'結論｜情境判斷句',R.judge);
  cmpT('文字 共識判斷句',SM+'差異｜判斷句',cv.judge); cmpT('文字 隱含倍數句',SM+'隱含｜隱含倍數句',cv.implied);
  // v4.4：年度差異原因（類型＋原因逐字；Excel 每個設定組合一列，未超過門檻時為空白）與原因摘要句
  cmpT('文字 差異原因摘要',SM+'差異｜差異原因摘要',cv.rsnSum);
  for(const k of Object.keys(X).filter(k=>k.startsWith(SM+'差異原因｜'))){ const [,yr,nm]=k.split('｜'), h=cv.rsn.find(x=>x.yr===yr&&x.name===nm), x=X[k];
    rows.push([(x[0]??'')===(h?h.type:'')&&(x[1]??'')===(h?h.text:'')?'OK ':'XX ','差異原因 '+yr+' '+nm,JSON.stringify(h?[h.type,h.text]:['','']),JSON.stringify([x[0],x[1]])]); }
  cv.rsn.forEach(h=>{ if(!X[SM+'差異原因｜'+h.yr+'｜'+h.name]) rows.push(['XX ','差異原因 '+h.yr+' '+h.name,'HTML 有原因、Excel 無此列（'+h.type+'）','']); });
  // v4.4：季度層（6 季模型／共識／指引／實際／差距／原因／句；焦點季；期間合計核對；摘要驗證點句）
  const qv=quarterlyView(d,p,q), QT='季度追蹤|', VSN={consensus:'vs 共識',guidance:'vs 指引中點',actual:'實際 vs 模型'};
  const cmpMix=(label,key,hs)=>{ const x=X[key]; if(!x){rows.push([label,'MISSING XL KEY '+key]);return;}
    const ok=hs.every((h,i)=>h==null?typeof x[i]!=='number':typeof h==='number'?typeof x[i]==='number'&&Math.abs(h-x[i])<1e-6:(x[i]??'')===h);
    rows.push([ok?'OK ':'XX ',label,JSON.stringify(hs.map(h=>typeof h==='number'?+h.toFixed(6):h)),JSON.stringify(x.slice(0,hs.length).map(v=>typeof v==='number'?+v.toFixed(6):v))]); };
  if(qv){ const n=qv.Qs.length;
    cmpMix('季度 平均在役MW',QT+'平均在役 MW（Billable，拆分依據）',qv.avgBil); cmpMix('季度 車隊折舊',QT+'車隊折舊（模型）',qv.da);
    for(const m of qv.rows){ const L=m.label, rq=m.q;
      cmpMix('季度 '+L+' 模型',QT+L+'（模型）',rq.map(r=>r.model));
      if(m.consensus) cmpMix('季度 '+L+' 共識',QT+'共識｜'+L+(m.consensus==='derived'?' [Derived]':''),rq.map(r=>r.cons));
      if(m.consensusSecondary) cmpMix('季度 '+L+' 次要共識',QT+(m.consensusSecondaryLabel||'共識｜'+m.consensusSecondary),rq.map(r=>r.consSec));
      if(rq.some(r=>r.guide)){ cmpMix('季度 '+L+' 指引低',QT+'指引｜'+L+'（低）',rq.map(r=>r.guide?r.guide[0]:null)); cmpMix('季度 '+L+' 指引高',QT+'指引｜'+L+'（高）',rq.map(r=>r.guide?r.guide[1]:null)); }
      cmpMix('季度 '+L+' 實際',QT+'實際｜'+(m.actualLabel||L),rq.map(r=>r.actual));
      for(const [g,t] of [['mc','模型 vs 共識'],['mg','模型 vs 指引中點'],['am','實際 vs 模型'],['ac','實際 vs 共識'],['ag','實際 vs 指引中點']]) cmpMix('季度 '+L+' '+t,QT+'差距｜'+L+'｜'+t,rq.map(r=>r[g]));
      if(m.marginOf) for(const [g,t] of [['mcP','模型 vs 共識'],['mgP','模型 vs 指引中點'],['amP','實際 vs 模型'],['acP','實際 vs 共識'],['agP','實際 vs 指引中點']]) cmpMix('季度 '+L+' '+t+' 利潤率pt',QT+'差距｜'+L+'｜'+t+'（利潤率 pt）',rq.map(r=>r[g]));
      if(rq.some(r=>r.gtext)) cmpMix('季度 '+L+' 文字指引',QT+'指引（文字）｜'+L,rq.map(r=>r.gtext?r.gtext.text:''));
      if(m.kind!=='pt') for(const vs of ['consensus','guidance','actual']) cmpMix('季度 '+L+' 原因 '+vs,QT+'原因｜'+L+'｜'+VSN[vs],rq.map(r=>{ const x=r.rsn.find(z=>z.vs===vs); return x?`${L} ${VSN[vs]} ${x.gtxt}：${x.type}（${x.text}）`:''; }));
      cmpMix('季度 '+L+' 句',QT+'句｜'+L,rq.map((r,i)=>qv.line(m,i)));
      const f=rq[qv.fi]; cmpMix('焦點 '+L,QT+'焦點｜'+L,[f.model,f.cons,f.mc,f.guide?f.guide[0]:(f.gtext?f.gtext.text:null),f.guide?f.guide[1]:null,f.mg,f.actual,f.am,f.ac,f.ag,f.mcP,f.mgP,f.amP,f.acP,f.agP]); }
    cmpMix('季度 差異原因一行',QT+'差異原因（一行）',qv.Qs.map((_,i)=>qv.rsnLine(i)));
    cmpMix('焦點 季度',QT+'焦點｜季度',[qv.focus.label]); cmpMix('焦點 差異原因',QT+'焦點｜差異原因',[qv.rsnLine(qv.fi)||'無（差距皆在 '+pctQ(qv.tol)+' 以內）']);
    for(const nm of ['營收','調整後 EBITDA','調整後營業利益','CapEx（毛額）','車隊折舊']) cmpMix('核對 '+nm,QT+'核對｜'+nm,qv.sums.map(()=>0));
    qv.key.forEach((m,i)=>cmpT('文字 驗證點 '+m.label,SM+'驗證｜'+m.label+'｜句',qv.keyLines[i])); } }
{ const G=evAnchorGrid(d,q,VAL_DEFAULTS); G.mults.forEach((m,i)=>{ cmp('矩陣腿 '+m+'x', V+'錨定×倍數｜EV/EBITDA 腿｜'+m.toFixed(1)+'x', G.leg[i]); cmp('矩陣目標 '+m+'x', V+'錨定×倍數｜加權目標價｜'+m.toFixed(1)+'x', G.tgt[i]); }); }
for(const r of rows) console.log(r.join(' | '));
// HTML P&L table FY26 display check
console.log('HTML P&L table FY26: 總營收(fy)',f[0].fyRevenue.toFixed(3),'營業利益(fy)',f[0].fyOpInc.toFixed(3),'營利率(2H)',f[0].opM,'EBITDA(2H)',f[0].ebitda.toFixed(3),'淨利率(2H)',f[0].nm.toFixed(3),'EPS(2H)',f[0].eps.toFixed(3));
console.log('HTML EPS all',f.map(e=>e.eps.toFixed(2)).join('/'),'ngEps',f.map(e=>e.ngEps.toFixed(2)).join('/'),'shares',f.map(e=>e.shares.toFixed(4)).join('/'));
console.log('ni',f.map(e=>e.ni.toFixed(3)).join('/'),'EBIT',f.map(e=>e.opInc.toFixed(2)).join('/'),'int',f.map(e=>e.interest.toFixed(2)).join('/'),'tax',f.map(e=>e.tax.toFixed(3)).join('/'));
rows.length=0;
const CO='評價_可比公司|';
for(const e of lM){ const key=e.labelXlsx;
  const xv=X[CO+key]; const ok=Math.abs(e.ev-xv[2])<0.02&&Math.abs(e.ev/e.rev-xv[4])<0.01&&Math.abs((e.ev+(e.opl||0))/e.rev-xv[5])<0.01&&Math.abs(e.ebitda-xv[6])<0.002;
  rows.push([ok?'OK ':'XX ','comps '+e.ticker,[e.ev.toFixed(2),(e.ev/e.rev).toFixed(2),((e.ev+(e.opl||0))/e.rev).toFixed(2)].join('/'),[xv[2],(+xv[4]).toFixed(2),(+xv[5]).toFixed(2)].join('/')]);}
const med=X[CO+'同業中位數']; rows.push([Math.abs(wM(lM.map(e=>e.ev/e.rev))-med[4])<0.01?'OK ':'XX ','comps median',wM(lM.map(e=>e.ev/e.rev)).toFixed(2),(+med[4]).toFixed(2)]);
const cr=X[CO+COMPANY_DATA.meta.ticker+'（TTM 至 Q2）']; const hv=(CALL_FACTS.mktCapLast+VAL_DEFAULTS.netDebt+CVN.reduce((a,c)=>a+c.M,0)+ndAdjQ(VAL_DEFAULTS))/CALL_FACTS.ttmRev; rows.push([Math.abs(hv-cr[4])<0.01?'OK ':'XX ','comps '+COMPANY_DATA.meta.ticker+' TTM EV/S',hv.toFixed(3),(+cr[4]).toFixed(3)]);
const cm=X[CO+T(COMPANY_DATA.meta.ticker+' 模型 «P1»E')]; const hv2=(VAL_DEFAULTS.price*p.shares+p.v.netDebt)/f[1].revenue; rows.push([Math.abs(hv2-cm[4])<0.01?'OK ':'XX ','comps '+COMPANY_DATA.meta.ticker+' FY27E EV/S',hv2.toFixed(3),(+cm[4]).toFixed(3)]);
for(const r of rows) console.log(r.join(' | '));
console.log('bench',d.totals.bench, 'X具名', JSON.stringify(X['站點租賃|具名站點合計']));
