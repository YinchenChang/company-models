// ===== segE.js：「總結」簡報頁（v3.4 新增）=====
// 所有數字即時取自資金引擎（dA）與評價引擎（TM），不另存常數；切換情境或修改輸入後自動更新。
// 版面固定為 1280×720 的投影片，一般檢視時等比縮放；簡報模式全螢幕逐頁播放；列印時每頁一張（16:9）。
// 全域名稱一律以 Q 結尾，避開模板函式庫。

var SW_Q = 1280, SH_Q = 720;

// v0.2：單價對照（company.json → priceCheck；只用於呈現，不進入計算）。實現單價＝最新一季營收 × 4 ÷ 在役 MW 估計
var PRICE_CHK_Q = COMPANY_DATA.priceCheck && COMPANY_DATA.priceCheck.inServiceMw > 0
  ? { ...COMPANY_DATA.priceCheck, realized: COMPANY_DATA.latestQuarter.revenue * 4 * 1e3 / COMPANY_DATA.priceCheck.inServiceMw } : null;

// v0.2：反向 DCF 在頁首與總結頁共用；輸入變動後延遲計算（約 0.4 秒），避免輸入時卡頓
function rvHookQ(e, o) {
  let [r, sr] = (0, v.useState)(null);
  (0, v.useEffect)(() => {
    let t = setTimeout(() => { try { sr(reverseDcf(e, o)); } catch { sr(null); } }, 30);
    return () => clearTimeout(t);
  }, [e, o]);
  return r;
}

// v0.2：副標題＝一句話答案（company.json → texts.subtitle 的佔位符以模型現值帶入，隨情境切換）
function prepayCoverQ(d) { let g = d.totals.gross; return g > 0 ? d.years.reduce((a, t) => a + t.prepayIn, 0) / g : NaN; }
function headlineQ(d, rv, e) {
  let c = prepayCoverQ(d), sc = e && [`low`, `base`, `high`].includes(e.scenario) ? e.scenario : `base`,
    kM = PMW_REVQ === `tkAnchor` ? d.m.revMW[4] * ((e && e.revScale) ?? 1) / tkAnchorQ(sc).anchor[4] : NaN, // v0.2a：模型每 MW 年收入 ÷ Tokenomics 錨（末期）
    vals = {
    prepayPct: Number.isFinite(c) ? hA(c * 100, 0) : `不適用`,
    gap: `$${Y(Math.max(0, -d.totals.preFinEnd), 1)}bn`,
    rvMult: rv ? (Number.isFinite(rv.Rt) ? Y(rv.Rt, 1) : `—`) : `…`,
    kMult: Number.isFinite(kM) ? Y(kM, 1) : `—`,
    rvAnchorMult: rv && Number.isFinite(rv.Rt) && Number.isFinite(kM) ? Y(rv.Rt * kM, 1) : `…`
  };
  return String(COMPANY_DATA.texts.subtitle || ``).replace(/\{(\w+)\}/g, (m, k) => vals[k] ?? m);
}

// 小工具：建立元素（children 為陣列時用 jsxs）
function elQ(tag, props, children) {
  let p = { ...(props || {}) };
  if (children !== void 0) p.children = children;
  return Array.isArray(children) ? (0, $.jsxs)(tag, p) : (0, $.jsx)(tag, p);
}

// 投影片在容器內等比縮放（所見即簡報與列印結果）
function ScaleQ({ children }) {
  let r = (0, v.useRef)(null), [w, sw] = (0, v.useState)(SW_Q);
  (0, v.useEffect)(() => {
    let n = r.current;
    if (!n) return;
    let f = () => sw(n.clientWidth || SW_Q);
    f();
    let o = typeof ResizeObserver < `u` ? new ResizeObserver(f) : null;
    o ? o.observe(n) : window.addEventListener(`resize`, f);
    return () => o ? o.disconnect() : window.removeEventListener(`resize`, f);
  }, []);
  let k = Math.min(1, w / SW_Q);
  return elQ(`div`, { ref: r, className: `sumQ-outer`, style: { width: `100%`, height: SH_Q * k } },
    elQ(`div`, { className: `sumQ-inner`, style: { width: SW_Q, height: SH_Q, transform: `scale(${k})`, transformOrigin: `top left` } }, children));
}

// v0.2：語意色（簡報各頁共用；圖例與文字同時標示，不只靠顏色）
var COLQ = { prepay: `#16a34a`, op: `#2563eb`, cash: `#94a3b8`, debt: `#0d9488`, conv: `#f59e0b`, eq: `#dc2626`, junk: `#7f1d1d`,
  model: `#0f766e`, cons: `#7c3aed`, real: `#dc2626`, need: `#111827`, off: `#9ca3af` };

// v0.2：內嵌 SVG 圖示（不外部載入）
var ICONQ = {
  bolt: `M13 2L4 14h7l-1 8 9-12h-7z`,
  server: `M4 4h16v6H4zM4 14h16v6H4zM8 7h.01M8 17h.01`,
  coin: `M12 3a9 9 0 1 0 0 18a9 9 0 1 0 0-18M15 9.5c-.6-1-1.7-1.5-3-1.5-1.7 0-3 .9-3 2s1.3 1.7 3 2 3 .9 3 2-1.3 2-3 2c-1.3 0-2.4-.5-3-1.5M12 6.5v11`,
  doc: `M6 3h9l4 4v14H6zM14 3v5h5M9 12h7M9 16h7`,
  chart: `M4 20V11M10 20V5M16 20v-7M2 20h20`,
  target: `M12 3a9 9 0 1 0 0 18a9 9 0 1 0 0-18M12 8a4 4 0 1 0 0 8a4 4 0 1 0 0-8M12 11.5v1`,
  eye: `M2 12s3.6-7 10-7 10 7 10 7-3.6 7-10 7S2 12 2 12zM12 9a3 3 0 1 0 0 6a3 3 0 1 0 0-6`,
  users: `M8 11a4 4 0 1 0 0-8a4 4 0 1 0 0 8M2 21v-1a6 6 0 0 1 12 0v1M17 11a3 3 0 1 0 0-6M16 15a5 5 0 0 1 6 5v1`,
  scale: `M12 3v18M5 7h14M5 7l-3 7a3 3 0 0 0 6 0zM19 7l-3 7a3 3 0 0 0 6 0zM8 21h8`
};
function IconQ({ n, size, color }) {
  return elQ(`svg`, { width: size || 28, height: size || 28, viewBox: `0 0 24 24`, fill: `none`, stroke: color || `currentColor`, strokeWidth: 1.9, strokeLinecap: `round`, strokeLinejoin: `round`, 'aria-hidden': `true`, style: { flex: `none` } },
    elQ(`path`, { d: ICONQ[n] || `` }));
}

// v0.2：大數字（≥56px）＋標籤
function BigQ({ label, value, sub, color, size, align, icon }) {
  return elQ(`div`, { className: `sumQ-a`, style: { textAlign: align || `left`, minWidth: 0 } }, [
    elQ(`div`, { key: `l`, style: { display: `flex`, alignItems: `center`, gap: 8, justifyContent: align === `center` ? `center` : `flex-start`, fontSize: 19, color: `var(--color-muted)`, fontWeight: 600 } },
      [icon ? elQ(IconQ, { key: `i`, n: icon, size: 24, color: color }) : null, label]),
    elQ(`div`, { key: `v`, style: { fontSize: size || 60, fontWeight: 800, lineHeight: 1.08, marginTop: 4, letterSpacing: `-.02em`, fontVariantNumeric: `tabular-nums`, color: color || `var(--color-fg)`, whiteSpace: `nowrap` } }, value),
    sub ? elQ(`div`, { key: `s`, style: { fontSize: 18, color: `var(--color-muted)`, marginTop: 6, lineHeight: 1.4 } }, sub) : null
  ]);
}

// v0.2：圖例
function LegQ({ items, size }) {
  return elQ(`div`, { style: { display: `flex`, flexWrap: `wrap`, gap: `8px 22px`, fontSize: size || 18 } }, items.map(([a, c]) =>
    elQ(`span`, { key: a, style: { display: `inline-flex`, alignItems: `center`, gap: 8 } }, [elQ(`i`, { key: `i`, style: { width: 16, height: 16, background: c, display: `inline-block`, borderRadius: 4 } }), a])));
}

// v0.2：頁底一句說明（≤ 約 40 字）
function NoteQ({ children }) {
  return elQ(`p`, { className: `sumQ-a sumQ-a4`, style: { fontSize: 18, color: `var(--color-muted)`, margin: `auto 0 0`, lineHeight: 1.5 } }, children);
}

// 單張投影片外框：主訊息標題＋內容＋頁尾
function SlideQ({ no, total, kicker, title, children, foot, tsize, main }) {
  return elQ(`section`, {
    className: `sumQ-slide`,
    style: {
      width: SW_Q, height: SH_Q, boxSizing: `border-box`, padding: `44px 56px 30px`,
      background: `var(--color-card)`, color: `var(--color-fg)`, border: `1px solid var(--color-border)`,
      borderRadius: 14, display: `flex`, flexDirection: `column`, overflow: `hidden`, fontFamily: `var(--font-sans)`
    }
  }, [
    elQ(`div`, { key: `k`, style: { fontSize: main ? 18 : 15, color: `var(--color-accent)`, fontWeight: 600, letterSpacing: main ? `.02em` : 0 } }, kicker),
    elQ(`h2`, { key: `t`, className: main ? `sumQ-a` : void 0, style: { margin: `8px 0 0`, fontSize: tsize || (main ? 38 : 34), lineHeight: 1.25, fontWeight: main ? 800 : 700, letterSpacing: `-.015em`, maxWidth: 1168, textWrap: `balance` } }, title),
    elQ(`div`, { key: `b`, style: { flex: 1, minHeight: 0, marginTop: main ? 22 : tsize ? 16 : 26, display: `flex`, flexDirection: `column` } }, children),
    elQ(`div`, { key: `f`, style: { display: `flex`, justifyContent: `space-between`, gap: 24, fontSize: 13, color: `var(--color-subtle)`, borderTop: `1px solid var(--color-border)`, paddingTop: 10 } }, [
      elQ(`span`, { key: `a` }, foot || ``),
      elQ(`span`, { key: `b`, style: { whiteSpace: `nowrap` } }, `${no} / ${total}`)
    ])
  ]);
}

// 水平長條（共同刻度）
function BarQ({ label, sub, val, max, color, neg, fmt }) {
  let pct = Math.max(0, Math.min(1, Math.abs(val) / (max || 1))) * 100;
  return elQ(`div`, { style: { display: `grid`, gridTemplateColumns: `330px 1fr 110px`, alignItems: `center`, gap: 16, padding: `9px 0` } }, [
    elQ(`div`, { key: `l` }, [
      elQ(`div`, { key: `a`, style: { fontSize: 19, fontWeight: 600 } }, label),
      sub ? elQ(`div`, { key: `b`, style: { fontSize: 13.5, color: `var(--color-muted)`, marginTop: 2 } }, sub) : null
    ]),
    elQ(`div`, { key: `b`, style: { height: 26, background: `var(--color-surface)`, borderRadius: 4 } },
      elQ(`div`, { style: { width: `${pct}%`, height: `100%`, background: color, borderRadius: 4, opacity: neg ? .55 : 1 } })),
    elQ(`div`, { key: `v`, style: { fontSize: 21, fontWeight: 700, textAlign: `right`, fontVariantNumeric: `tabular-nums`, color: neg ? `var(--color-bad)` : `var(--color-fg)` } }, (fmt || (x => (neg ? `−` : ``) + `$${Y(Math.abs(x), 1)}`))(val))
  ]);
}

// 大數字＋說明
function StatQ({ label, value, note, tone }) {
  return elQ(`div`, { style: { padding: `18px 20px`, background: `var(--color-surface)`, borderRadius: 10, borderLeft: `4px solid ${tone || `var(--color-accent)`}` } }, [
    elQ(`div`, { key: `l`, style: { fontSize: 15, color: `var(--color-muted)` } }, label),
    elQ(`div`, { key: `v`, style: { fontSize: 38, fontWeight: 700, marginTop: 4, fontVariantNumeric: `tabular-nums`, color: tone || `var(--color-fg)` } }, value),
    note ? elQ(`div`, { key: `n`, style: { fontSize: 14, color: `var(--color-muted)`, marginTop: 6, lineHeight: 1.5 } }, note) : null
  ]);
}


// v3.5：錨定年度 × 倍數矩陣（評價頁與總結頁共用）。big＝投影片字級
function EvGridQ({ st, o, g, big }) {
  let G0 = (0, v.useMemo)(() => g ? null : evAnchorGrid(runFunding(st), st, o), [g, st, o]), G = g || G0;
  let P = o.price, ck = o.evYear ?? 1, fs = big ? 17 : 12.5, pd = big ? `6px 12px` : `4px 8px`;
  let tab = (title, M, sub) => elQ(`div`, { key: title, style: { minWidth: 0 } }, [
    elQ(`div`, { key: `h`, style: { fontSize: fs + 1, fontWeight: 700, marginBottom: 6 } }, title),
    elQ(`table`, { key: `t`, style: { borderCollapse: `collapse`, width: `100%`, fontSize: fs, fontVariantNumeric: `tabular-nums` } }, [
      elQ(`thead`, { key: `h` }, elQ(`tr`, {}, [
        elQ(`th`, { key: `x`, style: { textAlign: `left`, padding: pd, borderBottom: `2px solid var(--color-ink)` } }, `倍數 ＼ 錨定`),
        ...G.years.map((y, j) => elQ(`th`, { key: y, style: { textAlign: `right`, padding: pd, borderBottom: `2px solid var(--color-ink)`, color: j + 1 === ck ? `var(--color-accent)` : `inherit` } }, y))
      ])),
      elQ(`tbody`, { key: `b` }, G.mults.map((m, i) => elQ(`tr`, { key: m }, [
        elQ(`td`, { key: `m`, style: { padding: pd, borderBottom: `1px solid var(--color-border)`, fontWeight: 600 } }, `${Y(m, 1)}x`),
        ...M[i].map((x, j) => {
          let on = j + 1 === ck && Math.abs(m - o.evEbitda) < 1e-9, hit = x >= P;
          return elQ(`td`, { key: j, style: { padding: pd, textAlign: `right`, borderBottom: `1px solid var(--color-border)`, background: on ? `#fff2a8` : hit ? `var(--color-ok-bg)` : `transparent`, fontWeight: on ? 700 : 400 } }, `$${Y(x, 1)}`);
        })
      ])))
    ]),
    sub ? elQ(`div`, { key: `s`, style: { fontSize: fs - 1.5, color: `var(--color-muted)`, marginTop: 6, lineHeight: 1.5 } }, sub) : null
  ]);
  return elQ(`div`, { style: { display: `grid`, gridTemplateColumns: `1fr 1fr`, gap: big ? 36 : 20, marginTop: big ? 0 : 8, marginBottom: big ? 0 : 20 } }, [
    tab(`加權目標價（每股）`, G.tgt, `DCF 腿 $${Y(G.dcf, 1)} × ${hA(G.wd * 100, 0)}＋EV/EBITDA 腿 × ${hA((1 - G.wd) * 100, 0)}`),
    tab(`EV/EBITDA 腿單獨（每股，折回 ${CALQ.targetText}）`, G.leg, `綠底＝不低於現價 $${Y(P, 2)}；折現率＝WACC ${hA(o.wacc * 100, 0)}`)
  ]);
}

// v4.3：共識資料的標記徽章（Verified／Interested-party）
function TagQ({ t }) {
  let ok = t === `Verified`;
  return elQ(`span`, { style: { fontSize: 11, fontWeight: 600, padding: `1px 6px`, borderRadius: 4, marginLeft: 6, whiteSpace: `nowrap`, verticalAlign: `middle`,
    background: ok ? `var(--color-ok-bg)` : `var(--color-surface)`, color: ok ? `var(--color-ok)` : `var(--color-watch)`, border: `1px solid var(--color-border)` } }, t);
}

// v4.3：一頁摘要（總結頁第 1 頁）的內容：2 × 2 四塊。所有數字依目前輸入與共識資料檔動態產生
function onePageQ({ cv, qv, TR, f, o, e, rv, scLabel, callTone }) {
  let C = CONSENSUS, PT = C.priceTarget, RA = C.ratings, AE = C.annualEstimates, QE = C.quarterlyEstimates,
    box = (k, title, kids) => elQ(`div`, { key: k, style: { background: `var(--color-surface)`, borderRadius: 10, padding: `8px 14px`, minHeight: 0, overflow: `hidden`, display: `flex`, flexDirection: `column` } }, [
      elQ(`div`, { key: `h`, style: { fontSize: 14.5, fontWeight: 700, color: `var(--color-accent)`, marginBottom: 4 } }, title), ...kids]),
    sm = { fontSize: 12, color: `var(--color-muted)`, lineHeight: 1.45 },
    td = (x, st) => elQ(`td`, { style: { padding: `1px 5px`, borderBottom: `1px solid var(--color-border)`, textAlign: `right`, whiteSpace: `nowrap`, ...(st || {}) } }, x),
    gcol = g => Math.abs(g) > CONS_TOL ? `var(--color-bad)` : `var(--color-fg)`,
    fmtV = (k, x) => k === `ebM` ? hA(x * 100, 1) : Y(x, 2),
    MET = [[`營收`, `rev`], [`調整後 EBITDA`, `ebitda`], [`EBITDA 率`, `ebM`], [`CapEx（毛額）`, `capex`], [`淨負債`, `nd`]],
    rvOk = rv && Number.isFinite(rv.R);
  let b1 = box(`b1`, `1｜結論`, [
    elQ(`div`, { key: `a`, style: { display: `flex`, alignItems: `baseline`, gap: 14 } }, [
      elQ(`span`, { key: `c`, style: { fontSize: 26, fontWeight: 800, color: callTone } }, f.call.call),
      elQ(`span`, { key: `p`, style: { fontSize: 26, fontWeight: 700, fontVariantNumeric: `tabular-nums` } }, `$${Y(TR.pt, 1)}`),
      elQ(`span`, { key: `u`, style: { fontSize: 15, color: `var(--color-muted)` } }, `現價 $${Y(o.price, 2)} · 空間 ${cv.up >= 0 ? `+` : `−`}${hA(Math.abs(cv.up) * 100, 0)}`)
    ]),
    elQ(`div`, { key: `g`, style: { display: `grid`, gridTemplateColumns: `repeat(3, 1fr)`, gap: 8, marginTop: 6 } }, [
      [`情境區間`, `$${Y(TR.A[0], 1)}–$${Y(TR.A[1], 1)}`], [TR.bLabel.replace(`（目前情境，`, `（`), `$${Y(TR.B[0], 1)}–$${Y(TR.B[1], 1)}`], [`賣出門檻（現價 −${pctQ(SELL_TH)}）`, `$${Y(TR.th, 1)}`]
    ].map(([a, b]) => elQ(`div`, { key: a, style: { background: `var(--color-card)`, borderRadius: 8, padding: `6px 10px` } }, [
      elQ(`div`, { key: `a`, style: { fontSize: 12, color: `var(--color-muted)` } }, a),
      elQ(`div`, { key: `b`, style: { fontSize: 18, fontWeight: 700, fontVariantNumeric: `tabular-nums` } }, b)]))),
    elQ(`p`, { key: `h`, style: { fontSize: 13.5, lineHeight: 1.45, margin: `6px 0 0` } }, cv.head),
    elQ(`p`, { key: `j`, style: { ...sm, fontSize: 11.5, margin: `3px 0 0` } }, TR.judge)
  ]);
  let b2 = box(`b2`, `2｜與市場的差異（模型：${scLabel} vs 共識）`, [
    elQ(`table`, { key: `t`, style: { borderCollapse: `collapse`, width: `100%`, fontSize: 12, fontVariantNumeric: `tabular-nums` } }, [
      elQ(`thead`, { key: `h` }, elQ(`tr`, {}, [td(`US$bn`, { textAlign: `left`, fontWeight: 600, borderBottom: `2px solid var(--color-ink)` }),
        ...CONS_YEARS.map(yr => td(`${yr} 模型／共識／差距`, { fontWeight: 600, borderBottom: `2px solid var(--color-ink)` }))])),
      elQ(`tbody`, { key: `b` }, MET.map(([n, k]) => elQ(`tr`, { key: k }, [td(n, { textAlign: `left` }),
        ...cv.rows.map(r => td([fmtV(k, r.m[k]), `／`, fmtV(k, r.c[k]), `　`, elQ(`b`, { key: `g`, style: { color: k === `nd` || k === `ebM` ? `var(--color-fg)` : gcol(r.gap[k]) } }, k === `ebitda` || k === `nd` ? dTxtQ(r.dif[k], `US$bn`) : gapTxtQ(r.gap[k], k === `ebM`))]))])))
    ]),
    elQ(`p`, { key: `j`, style: { fontSize: 13, lineHeight: 1.4, margin: `5px 0 0`, fontWeight: 600 } }, cv.judge),
    cv.rsnSum ? elQ(`p`, { key: `r`, style: { fontSize: 11.5, lineHeight: 1.35, margin: `2px 0 0` } }, cv.rsnSum) : null
  ]);
  let b3 = box(`b3`, `3｜現價隱含什麼`, [
    elQ(`div`, { key: `g`, style: { display: `grid`, gridTemplateColumns: `repeat(3, 1fr)`, gap: 8 } }, [
      [`共識平均目標價隱含`, `${Y(cv.impTgt, 1)}x`, `$${Y(PT.mean, 2)}，股數 ${Y(o.shares, 3)}bn`],
      [`現價隱含`, `${Y(cv.impPx, 1)}x`, `$${Y(o.price, 2)}；同一共識 FY28 數字`],
      [`模型方法區間上緣`, `${multTxt(cv.mHi)}x`, `錨定 ${PERIOD_LABELS[o.evYear ?? 1]}、折回 ${CALQ.targetText}`]
    ].map(([a, b, c]) => elQ(`div`, { key: a, style: { background: `var(--color-card)`, borderRadius: 8, padding: `6px 10px` } }, [
      elQ(`div`, { key: `a`, style: { fontSize: 12, color: `var(--color-muted)` } }, a),
      elQ(`div`, { key: `b`, style: { fontSize: 19, fontWeight: 700, fontVariantNumeric: `tabular-nums` } }, b),
      elQ(`div`, { key: `c`, style: { fontSize: 10.5, color: `var(--color-muted)`, lineHeight: 1.35 } }, c)]))),
    elQ(`p`, { key: `i`, style: { fontSize: 13, lineHeight: 1.4, margin: `5px 0 0`, fontWeight: 600 } }, cv.implied),
    elQ(`p`, { key: `r`, style: { fontSize: 12.5, lineHeight: 1.4, margin: `4px 0 0` } }, rvOk
      ? `反向 DCF（DCF＝現價，其他不變）：FY30 每 MW 年收入需 $${Y(rv.rev30 * rv.R, 1)}m（${rv.R >= 1 ? `+` : `−`}${hA(Math.abs(rv.R - 1) * 100, 0)}），或建置成本 ${Number.isFinite(rv.C) ? `$${Y(rv.cost30 * rv.C, 1)}m（${rv.C >= 1 ? `+` : `−`}${hA(Math.abs(rv.C - 1) * 100, 0)}）` : `無解（降到 1 成仍不夠）`}，或穩態 EBITDA 率 ${Number.isFinite(rv.Eb) ? hA(rv.Eb * 100, 0) : `無解`}${Number.isFinite(rv.Rt) ? `；加權目標價＝現價需每 MW 年收入 ${rv.Rt >= 1 ? `+` : `−`}${hA(Math.abs(rv.Rt - 1) * 100, 0)}` : ``}。`
      : `反向 DCF：計算中或無解。`),
    elQ(`p`, { key: `n`, style: { ...sm, margin: `3px 0 0`, fontSize: 11 } }, `隱含倍數＝（價格 × 股數＋共識 FY28 淨負債）÷ 共識 FY28 調整後 EBITDA。`)
  ]);
  let RL = [[`強力買進`, RA.strongBuy], [`買進`, RA.buy], [`持有`, RA.hold], [`賣出`, RA.sell], [`強力賣出`, RA.strongSell]];
  let b4 = box(`b4`, `4｜驗證點與市場看法`, [
    ...(qv ? [elQ(`div`, { key: `q`, style: { fontSize: 13, fontWeight: 600 } }, [`${qv.focus.label} 財報${qv.focus.reportNote ? `（${qv.focus.reportNote}）` : ``}：模型｜共識｜指引`, QE ? elQ(TagQ, { key: `t`, t: QE.tag }) : null]),
      ...qv.keyLines.map((t, i) => elQ(`div`, { key: `k${i}`, style: { fontSize: 12, lineHeight: 1.35, marginTop: 2 } }, t))] : []),
    elQ(`p`, { key: `r`, style: { fontSize: 12.5, lineHeight: 1.4, margin: `4px 0 0` } }, [
      `評等分布（${RA.month}，${RA.total} 家）：${RL.map(([a, b]) => `${a} ${b}`).join(`／`)}；目標價平均 $${Y(PT.mean, 2)}、中位數 $${Y(PT.median, 2)}、區間 $${Y(PT.low, 0)}–$${Y(PT.high, 0)}（${PT.analysts} 家）`, elQ(TagQ, { key: `t`, t: PT.tag })])
  ]);
  return [
    elQ(`div`, { key: `g`, style: { flex: 1, minHeight: 0, display: `grid`, gridTemplateColumns: `1fr 1.22fr`, gridTemplateRows: `1fr 1fr`, gap: 8 } }, [b1, b2, b3, b4]),
    elQ(`p`, { key: `s`, style: { fontSize: 10.5, color: `var(--color-muted)`, lineHeight: 1.4, margin: `6px 0 0` } },
      `共識來源：${AE.source}、${PT.source}；擷取 ${AE.retrieved}。來源獨立性：${C.sourceIndependence}`)
  ];
}

// v4.3：「損益與評價 → 市場共識」分頁：模型 vs 共識對照、共識資料逐筆（數值、來源、擷取日期、標記）、來源獨立性與未取得清單
function ConsTabQ({ d, p, o, tr: TR, st }) {
  let cv = (0, v.useMemo)(() => consensusView(d, p, o, TR, st), [d, p, o, TR, st]), items = (0, v.useMemo)(() => consensusItems(), []),
    C = CONSENSUS, secs = [...new Set(items.map(x => x.sec))],
    th = (x, r) => elQ(`th`, { key: x, className: `px-2 py-1.5 text-xs font-semibold text-muted ${r ? `text-right` : `text-left`}` }, x),
    fmtI = x => x.vals.length ? x.vals.map(z => x.unit === `%` ? hA(z * 100, x.dp) : `${x.unit === `US$` ? `$` : ``}${Y(z, x.dp)}`).join(`／`) : x.text,
    MET = [[`營收`, `rev`], [`調整後 EBITDA`, `ebitda`], [`EBITDA 率`, `ebM`], [`CapEx（毛額）`, `capex`], [`淨負債`, `nd`]];
  return elQ(Jj, { className: `space-y-4` }, [
    elQ(`div`, { key: `h` }, [
      elQ(`h2`, { key: `t`, className: `font-display text-lg font-semibold` }, `市場共識`),
      elQ(`p`, { key: `p`, className: `mt-1 text-xs leading-relaxed text-muted` },
        `資料檔 ${COMPANY_DATA.meta.consensusFile}（${C.asOf}；只讀，不增補數字）。標記：Verified＝${C.tagLegend.Verified}；Interested-party＝${C.tagLegend[`Interested-party`]}。`)
    ]),
    elQ(`div`, { key: `si`, className: `rounded-lg border border-watch/40 bg-watch/5 px-3 py-2 text-sm leading-relaxed` }, [elQ(`b`, { key: `b` }, `來源獨立性：`), C.sourceIndependence]),
    elQ(`div`, { key: `cmp`, className: `max-w-full overflow-x-auto` }, [
      elQ(`h3`, { key: `h`, className: `text-sm font-semibold` }, `模型（目前情境）vs 共識：FY26–FY28`),
      elQ(`table`, { key: `t`, className: `mt-2 w-full text-sm` }, [
        elQ(`thead`, { key: `h` }, elQ(`tr`, { className: `border-b border-border` }, [th(`US$bn`), ...CONS_YEARS.flatMap(y => [th(`${y} 模型`, 1), th(`共識`, 1), th(`差距`, 1)])])),
        elQ(`tbody`, { key: `b` }, MET.map(([n, k]) => elQ(`tr`, { key: k, className: `border-t border-border` }, [
          elQ(`td`, { key: `n`, className: `px-2 py-1.5 font-medium` }, n),
          ...cv.rows.flatMap((r, i) => [r.m[k], r.c[k], r.gap[k]].map((x, j) => elQ(`td`, {
            key: `${i}${j}`, className: `px-2 py-1.5 text-right font-mono tabular-nums ${j === 2 && k !== `ebM` && k !== `nd` && Math.abs(x) > CONS_TOL ? `text-bad font-semibold` : ``}`
          }, j === 2 ? (k === `ebitda` || k === `nd` ? dTxtQ(r.dif[k], `US$bn`) : gapTxtQ(x, k === `ebM`)) : k === `ebM` ? hA(x * 100, 1) : Y(x, 2))))
        ])))
      ]),
      elQ(`p`, { key: `j`, className: `mt-2 text-sm font-semibold` }, cv.judge),
      cv.rsn.length ? elQ(`div`, { key: `r`, className: `mt-1 text-sm` }, [elQ(`div`, { key: `h`, className: `font-semibold` }, cv.rsnSum.replace(/；[^；]*$/, ``)),
        ...cv.rsn.map((x, i) => elQ(`div`, { key: i, className: `text-xs leading-relaxed` }, `${x.yr} ${x.name} ${x.gtxt}｜${x.type}：${x.text}`))]) : null,
      elQ(`p`, { key: `i`, className: `mt-1 text-sm` }, cv.implied),
      elQ(`p`, { key: `n`, className: `mt-1 text-xs leading-relaxed text-muted` },
        `判斷只用營收、EBITDA、CapEx，門檻 ${pctQ(CONS_TOL)}（company.json → methodology.consensusGapTol）；營收、CapEx 以比例列差距，調整後 EBITDA 與淨負債以金額差、EBITDA 率以百分點列示（判斷門檻仍以比例計）；淨負債只列不判斷（共識口徑未揭露）。FY26 模型為全年口徑：營收＝1H 實際＋2H 模型；調整後 EBITDA＝1H 實際 ${Y(ACTUAL_1H.adjEbitda, 3)}（Q1 ${Y(ACTUAL_1H.adjEbitdaMeta.q1, 3)}＋Q2 ${Y(ACTUAL_1H.adjEbitdaMeta.q2, 3)}；${ACTUAL_1H.adjEbitdaMeta.tag}：${ACTUAL_1H.adjEbitdaMeta.note}；${ACTUAL_1H.adjEbitdaMeta.sources.map(x => `${x.quarter} ${x.name}`).join(`、`)}；${ACTUAL_1H.adjEbitdaMeta.crossCheck}）＋2H 模型，只用於本對照，不改模型 GAAP 損益與評價；CapEx＝1H 實際毛額 ${Y(ACTUAL_1H.capex, 3)}＋2H 模型毛額（共識口徑未揭露）。隱含倍數＝（價格 × 股數 ${Y(o.shares, 3)}bn＋共識 FY28 淨負債）÷ 共識 FY28 調整後 EBITDA。`)
    ]),
    ...secs.map(sc => elQ(`div`, { key: sc, className: `max-w-full overflow-x-auto` }, [
      elQ(`h3`, { key: `h`, className: `text-sm font-semibold` }, sc),
      elQ(`table`, { key: `t`, className: `mt-1 w-full text-sm` }, [
        elQ(`thead`, { key: `h` }, elQ(`tr`, { className: `border-b border-border` }, [th(`項目`), th(`數值`, 1), th(`來源`), th(`擷取日期`), th(`標記`), th(`備註`)])),
        elQ(`tbody`, { key: `b` }, items.filter(x => x.sec === sc).map(x => elQ(`tr`, { key: x.label, className: `border-t border-border align-top` }, [
          elQ(`td`, { key: `a`, className: `px-2 py-1.5 font-medium` }, x.label.replace(/^共識｜/, ``)),
          elQ(`td`, { key: `b`, className: `px-2 py-1.5 ${x.vals.length ? `text-right font-mono tabular-nums whitespace-nowrap` : `text-xs`}` }, fmtI(x)),
          elQ(`td`, { key: `c`, className: `px-2 py-1.5 text-xs` }, x.url ? elQ(`a`, { href: x.url, target: `_blank`, rel: `noreferrer`, className: `underline` }, x.src) : x.src),
          elQ(`td`, { key: `d`, className: `px-2 py-1.5 text-xs whitespace-nowrap` }, x.date),
          elQ(`td`, { key: `e`, className: `px-2 py-1.5 text-xs whitespace-nowrap` }, x.tag),
          elQ(`td`, { key: `f`, className: `px-2 py-1.5 text-xs text-muted` }, x.note)
        ])))
      ])
    ]))
  ]);
}

// v4.4：「資金模型 → 各期收支 → 季度追蹤」分頁。第一屏：焦點季各指標的模型／共識／指引／實際與差距＋一行差異原因；
// 6 季路徑、期間合計核對、拆分方法放次層（預設收合）。全部由 quarterlyView 產生，缺資料顯示「不適用」。
function QuarterTabQ({ d, p, st }) {
  let qv = (0, v.useMemo)(() => quarterlyView(d, p, st), [d, p, st]), [sel, setSel] = (0, v.useState)(null);
  if (!qv) return elQ(`p`, { className: `text-sm text-muted` }, `company.json 未設定 quarterly：季度層不適用。`);
  let i = sel ?? qv.fi, Q = qv.Qs[i], C = QTR_CFG,
    th = (x, k) => elQ(`th`, { key: k ?? x, style: { ...xstyQ.th, whiteSpace: `nowrap` } }, x),
    thL = x => elQ(`th`, { key: x, style: xstyQ.thL }, x),
    td = (x, k, st2) => elQ(`td`, { key: k, style: { ...xstyQ.td, ...(st2 || {}) } }, x),
    tdL = (x, k) => elQ(`td`, { key: k, style: xstyQ.tdL }, x),
    gp = (r, k, m) => r[k] == null ? NA : gTxtQ(r[k], m.kind, m.unit, r[k + `P`]),
    gcol = (r, vs) => r.rsn.some(x => x.vs === vs) ? { color: `#9f1239`, fontWeight: 700 } : {},
    vt = (x, u, miss) => x == null ? miss : valTxtQ(x, u),
    NA = `不適用`, TBD = `待公布`,
    rl = qv.rsnLine(i),
    tbl = (head, body, minW) => elQ(`div`, { style: xstyQ.wrap }, elQ(`table`, { style: { ...xstyQ.table, minWidth: minW || 760 } }, [
      elQ(`thead`, { key: `h` }, elQ(`tr`, {}, head)), elQ(`tbody`, { key: `b` }, body)]));
  let first = tbl([thL(`指標`), ...[`模型`, `共識`, `模型 vs 共識`, `指引`, `模型 vs 指引中點`, `實際`, `實際 vs 模型`, `實際 vs 共識`, `實際 vs 指引`].map(x => th(x))],
    qv.rows.map(m => { let r = m.q[i], u = m.unit, hasA = r.actual != null;
      return elQ(`tr`, { key: m.key }, [tdL(m.label, `l`), td(vt(r.model, u, NA), `m`), td(vt(r.cons, u, NA), `c`), td(gp(r, `mc`, m), `mc`, gcol(r, `consensus`)),
        td(r.guide ? rngTxtQ(r.guide, u) : r.gtext ? `${r.gtext.text}（文字）` : NA, `g`), td(r.guide ? `${gp(r, `mg`, m)}（${r.posM}）` : NA, `mg`, gcol(r, `guidance`)),
        td(hasA ? valTxtQ(r.actual, u) : TBD, `a`), td(hasA ? gp(r, `am`, m) : `—`, `am`, gcol(r, `actual`)), td(hasA ? gp(r, `ac`, m) : `—`, `ac`),
        td(hasA ? (r.guide ? `${gp(r, `ag`, m)}（${r.posA}）` : NA) : `—`, `ag`)]) }), 980);
  let L = qv.Qs.map(q => q.label),
    path = tbl([thL(`項目`), th(`單位`), ...L.map((x, k) => th(x, k))], [
      ...qv.rows.map(m => [m.label + `（模型）`, m.unit, m.q.map(r => r.model)]),
      [`車隊折舊（模型）`, `US$bn`, qv.da], [`平均在役 MW（Billable，拆分依據）`, `MW`, qv.avgBil],
      ...qv.rows.filter(m => m.consensus).map(m => [`共識｜${m.label}${m.consensus === `derived` ? ` [Derived]` : ``}`, m.unit, m.q.map(r => r.cons)]),
      ...qv.rows.filter(m => m.consensusSecondary).map(m => [m.consensusSecondaryLabel || `共識｜${m.consensusSecondary}`, m.unit, m.q.map(r => r.consSec)]),
      ...qv.rows.filter(m => m.q.some(r => r.guide || r.gtext)).map(m => [`指引｜${m.label}`, m.unit, m.q.map(r => r.guide || r.gtext), 1]),
      ...qv.rows.map(m => [`實際｜${m.actualLabel || m.label}`, m.unit, m.q.map(r => r.actual), 2])
    ].map(([a, u, xs, kind], k) => elQ(`tr`, { key: k }, [tdL(a, `a`), td(u, `u`, { textAlign: `left`, color: `#6b7280` }),
      ...xs.map((x, j) => td(kind === 1 ? (!x ? NA : Array.isArray(x) ? rngTxtQ(x, u) : `${x.text}（文字，${x.tag}）`) : x == null ? (kind === 2 ? TBD : NA) : valTxtQ(x, u).replace(/bn$|\sMW$/, ``).replace(`$`, ``), j))])));
  let sums = tbl([thL(`期間`), th(`指標`), th(`季度加總`), th(`年度模型`), th(`差額`)],
    qv.sums.flatMap(s => [[`revenue`, `營收`], [`adjEbitda`, `調整後 EBITDA`], [`adjOpInc`, `調整後營業利益`], [`capex`, `CapEx（毛額）`], [`da`, `車隊折舊`]].map(([k, n]) =>
      elQ(`tr`, { key: s.name + k }, [tdL(s.name, `p`), td(n, `n`, { textAlign: `left` }), td(Y(s[k].q, 3), `q`), td(Y(s[k].a, 3), `a`), td(Y(s[k].q - s[k].a, 6), `d`)]))));
  let HOW = { anchor: `從${C.revenueAnchor ? ` ${C.revenueAnchor.label} ${valTxtQ(C.revenueAnchor.value, `US$bn`)}` : `最新一季實際`}起逐季線性爬升`, driverAvg: qv.drv ? `依平均在役 MW 分配` : `平分（無 MW 資料）`, equal: `平分` },
    meth = [`營收：${C.periodNames.map((n, j) => `${n} ${HOW[(C.revenueSplit || [])[j] || `equal`]}`).join(`；`)}。`,
      `EBITDA 率：跨季一條直線，同時符合各期年度 EBITDA。調整後營業利益＝調整後 EBITDA − 車隊折舊（依期內平均 PP&E）。`,
      `CapEx：${C.capexSplit === `guidanceAnchor` ? `有季度指引的季取指引中點，其餘季分配期間餘數（` : ``}${qv.drv ? `依季內新增 Accepted MW` : `期內平分`}${C.capexSplit === `guidanceAnchor` ? `）` : ``}。${qv.drv && C.driver.endStartNote ? `MW 起點：${C.driver.endStartNote}。` : ``}`,
      `差距：營收、CapEx 為比例；利潤類為金額差／利潤率百分點；EBITDA 率為百分點；MW 為差額。需附原因：營收、CapEx、MW 差距超過比較基準的 ${pctQ(qv.tol)}，利潤類利潤率差超過 ${PM_TOL != null ? `${Y(PM_TOL * 100, 1)}pt` : pctQ(qv.tol)}，或模型落在指引區間外。有期間合計時拆解為水準與分配：分配較大為「拆法」、水準較大為「觀點」；其餘讀 company.json → varianceReasons。季度層只用於追蹤，不影響年度數字與目標價。`];
  return elQ(`div`, { className: `space-y-3` }, [
    elQ(hdrQ, { key: `h`, title: `季度追蹤（${Q.label}${Q.reportNote ? `，財報${Q.reportNote}` : ``}）`, tip: `由年度模型拆分（目前情境），季度加總＝年度；只用於追蹤，不影響目標價。實際數讀 company.json → quarterly.actuals。` }),
    elQ(`div`, { key: `s`, className: `flex flex-wrap gap-1` }, qv.Qs.map((q, k) => elQ(`button`, { key: q.key, onClick: () => setSel(k),
      className: `h-8 rounded-md px-3 text-xs font-medium ${k === i ? `bg-ink text-accent-fg` : `bg-surface text-muted hover:text-fg`}` }, q.label))),
    elQ(`div`, { key: `t` }, first),
    elQ(`p`, { key: `r`, className: `text-sm leading-relaxed` }, [elQ(`b`, { key: `b` }, `差異原因：`), rl || `無（差距皆在 ${pctQ(qv.tol)} 以內）`]),
    elQ(accQ, { key: `a1`, title: `6 季路徑（模型、共識、指引、實際）`, sum: `${L[0]}–${L[L.length - 1]}` }, path),
    elQ(accQ, { key: `a2`, title: `期間合計核對（季度加總＝年度模型）`, sum: qv.sums.map(s => s.name).join(`／`) }, sums),
    elQ(accQ, { key: `a3`, title: `拆分方法`, sum: `` }, meth.map((t, k) => elQ(`p`, { key: k, className: `text-sm leading-relaxed` }, t)))
  ]);
}

var SUMCSSQ = `
@media print {
  @page { size: 1280px 720px; margin: 0; }
  html, body { background: #fff !important; }
  .sumQ-np, .sumQ-toolbar, .sumQ-pres { display: none !important; }
  .min-h-screen { min-height: 0 !important; }
  .sumQ-main { padding: 0 !important; margin: 0 !important; max-width: none !important; }
  .sumQ-main > * { margin: 0 !important; }
  .sumQ-wrap { padding: 0 !important; margin: 0 !important; gap: 0 !important; display: block !important; }
  .sumQ-outer { height: 720px !important; width: 1280px !important; overflow: hidden; }
  .sumQ-outer + .sumQ-outer { break-before: page; page-break-before: always; }
  .sumQ-inner { transform: none !important; }
  .sumQ-slide { border: 0 !important; border-radius: 0 !important; }
  * { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
}
/* v0.2：簡報模式換頁時的進場動畫（只在 .sumQ-pres 內；一般檢視、列印與 PDF 不播放，數字不會閃爍或遺漏） */
@keyframes sumQin { from { opacity: 0; transform: translateY(14px); } to { opacity: 1; transform: none; } }
@keyframes sumQgx { from { transform: scaleX(0); } to { transform: scaleX(1); } }
@keyframes sumQgy { from { transform: scaleY(0); } to { transform: scaleY(1); } }
.sumQ-pres .sumQ-a { animation: sumQin .5s ease-out both; }
.sumQ-pres .sumQ-a2 { animation-delay: .12s; }
.sumQ-pres .sumQ-a3 { animation-delay: .24s; }
.sumQ-pres .sumQ-a4 { animation-delay: .4s; }
.sumQ-pres .sumQ-g { animation: sumQgx .8s cubic-bezier(.2,.7,.2,1) .25s both; transform-origin: left center; }
.sumQ-pres .sumQ-gy { animation: sumQgy .8s cubic-bezier(.2,.7,.2,1) .25s both; transform-origin: center bottom; }
@media (prefers-reduced-motion: reduce) { .sumQ-pres * { animation: none !important; } }
@media print { .sumQ-a, .sumQ-g, .sumQ-gy { animation: none !important; opacity: 1 !important; transform: none !important; } }
.sumQ-btn { height: 40px; padding: 0 16px; border-radius: 8px; font-size: 14px; font-weight: 600; cursor: pointer; border: 1px solid var(--color-border); background: var(--color-card); color: var(--color-fg); }
.sumQ-btn:hover { background: var(--color-surface); }
.sumQ-btn:focus-visible { outline: 2px solid var(--color-accent); outline-offset: 2px; }
.sumQ-btn.pri { background: var(--color-ink); color: var(--color-accent-fg); border-color: var(--color-ink); }
`;

function SumQ({ d, f, e, o, m, tr: TR, active, rv }) {
  // --- 情境比較（保留目前手動調整，只換擴張力道）---
  let scen = (0, v.useMemo)(() => [`low`, `base`, `high`].map(sc => {
    let st = scnQ(e, sc), dd = runFunding(st), pp = runValuation(dd, st, o), po = runValuation(dd, st, { ...o, dcfMode: `option` });
    return {
      sc, label: SCENARIOS[sc].label,
      capex: dd.years.reduce((a, y) => a + y.gross, 0),
      gap: Math.max(0, -dd.totals.preFinEnd),
      nd: dd.totals.newDebt, cv: dd.totals.convNew, eq: dd.totals.equity, jk: dd.totals.junk,
      debt30: dd.years[4].totalDebtEnd,
      rev30: pp.fwd[4].fyRevenue, eb30: pp.fwd[4].fyEbitda,
      dcf: pp.d.invalid ? NaN : pp.d.perShareT, ev: pp.peAdj,
      tgt: pp.call.blended, opt: po.call.blended, call: pp.call.call
    };
  }), [e, o]);
  // --- 反向 DCF：由頁首（segD）以 rvHookQ 計算後傳入（v0.2）---
  // --- 預付款對照（v0.2）：同一引擎，只把客戶預付比率設為 0（呈現用試算，不改模型數字）---
  let p0 = (0, v.useMemo)(() => {
    let st0 = { ...e, a: { ...e.a, customerFund: e.a.customerFund.map(() => 0) } }, d0 = runFunding(st0), v0 = runValuation(d0, st0, o);
    return { d: d0, tgt: v0.call.blended };
  }, [e, o]);
  let cv = (0, v.useMemo)(() => consensusView(d, f, o, TR, e), [d, f, o, TR, e]), qv = (0, v.useMemo)(() => quarterlyView(d, f, e), [d, f, e]);
  let eg = (0, v.useMemo)(() => evAnchorGrid(d, e, o), [d, e, o]);
  let km = (0, v.useMemo)(() => PMW_REVQ === `tkAnchor` ? tkSensQ(e, o, !0).matrix : null, [e, o]); // v0.2a：容量 × 價格 3 × 3（目前輸入，同一引擎）
  let [pres, sp] = (0, v.useState)(!1), [ix, si] = (0, v.useState)(0), pr = (0, v.useRef)(null);

  let y = d.years, T = d.totals, b = rpoBridge(d), P = o.price, tgt = f.call.blended, up = f.call.upside;
  let scLabel = e.scenario === `custom` ? `自訂情境` : SCENARIOS[e.scenario]?.label;
  let ver = VLOG[VLOG.length - 1][0];
  let foot = `${COMPANY_DATA.texts.title} · ${COMPANY_DATA.meta.company} 收支模型 ${ver}（${UPDATE_DATE}）· ${scLabel} · 現價 $${Y(P, 2)}（${COMPANY_DATA.meta.priceDate} 收盤）· 研究框架，不是投資建議`;
  let callTone = f.call.call === `買進` ? `var(--color-ok)` : f.call.call === `賣出` ? `var(--color-bad)` : `var(--color-watch)`;

  // 單位經濟（與頁首摘要同一公式）
  let rev30 = d.m.revMW[4] * (e.revScale ?? 1) * 1e3, util30 = d.m.util[4] / 100, eb30 = y[4].ebM;
  let ebMW = rev30 * util30 * eb30, cost30 = e.a.costMW[4] * (e.capexScale ?? 1);
  let crf = o.wacc / (1 - Math.pow(1 + o.wacc, -e.gpuLife)), recov = cost30 * crf;
  let beRev = recov / (util30 * eb30), beCost = ebMW / crf;

  // 模型限制用數字
  let rentGap = y.reduce((a, t) => a + (t.rentBench - t.lease), 0);
  let eqYears = y.filter(t => t.equity > .05).map(t => t.year);
  let rateLo = Math.min(...e.m.rate), rateHi = Math.max(...e.m.rate);

  let slides = [];
  // v0.2：正式簡報（一頁一個訊息、以圖與大數字為主）＋附錄（既有明細頁）＋最後一頁一頁摘要
  let M = (kicker, title, body, tsize) => slides.push({ kicker, title, body, tsize, main: !0 });
  let S = (kicker, title, body, tsize) => slides.push({ kicker: `附錄｜${kicker}`, title, body, tsize });
  let sumY = k => y.reduce((a, t) => a + (t[k] || 0), 0), money = (x, n) => `$${Y(x, n ?? 1)}bn`;
  let preIn = sumY(`prepayIn`), cover = prepayCoverQ(d), gap0 = Math.max(0, -T.preFinEnd), shM = x => `${Y(x * 1e3, 0)}m`;
  let callCol = c => c === `買進` ? `var(--color-ok)` : c === `賣出` ? `var(--color-bad)` : `var(--color-watch)`;
  let curSc = e.scenario === `custom` ? null : e.scenario;

  // 1｜封面兼結論：主標題＋一句答案＋三情境目標價 vs 現價
  {
    let hi = Math.max(P, ...scen.map(s => s.tgt), TR.th) * 1.08, X = x => `${Math.max(0, Math.min(100, x / hi * 100))}%`;
    M(`${COMPANY_DATA.meta.company}（${COMPANY_DATA.meta.ticker}）收支模型 · ${ver} · ${UPDATE_DATE}`, COMPANY_DATA.texts.title, [
      elQ(`p`, { key: `a`, className: `sumQ-a sumQ-a2`, style: { fontSize: 25, lineHeight: 1.45, margin: 0, fontWeight: 600, color: `var(--color-accent)`, maxWidth: 1120 } }, headlineQ(d, rv, e)),
      elQ(`div`, { key: `g`, style: { display: `grid`, gridTemplateColumns: `1.05fr 1fr 1fr 1fr`, gap: 18, marginTop: 26 } }, [
        elQ(`div`, { key: `p`, className: `sumQ-a sumQ-a2`, style: { padding: `16px 20px`, borderRadius: 12, background: `var(--color-ink)`, color: `var(--color-accent-fg)` } }, [
          elQ(`div`, { key: `l`, style: { fontSize: 19, opacity: .8, fontWeight: 600 } }, `現價（${COMPANY_DATA.meta.priceDate}）`),
          elQ(`div`, { key: `v`, style: { fontSize: 60, fontWeight: 800, lineHeight: 1.1, fontVariantNumeric: `tabular-nums` } }, `$${Y(P, 2)}`),
          elQ(`div`, { key: `s`, style: { fontSize: 18, opacity: .8, marginTop: 4 } }, `賣出門檻 $${Y(TR.th, 1)}`)
        ]),
        ...scen.map((s, i) => elQ(`div`, { key: s.sc, className: `sumQ-a sumQ-a3`, style: { padding: `16px 20px`, borderRadius: 12, background: `var(--color-surface)`, border: s.sc === curSc ? `2px solid var(--color-accent)` : `2px solid transparent` } }, [
          elQ(`div`, { key: `l`, style: { fontSize: 19, color: `var(--color-muted)`, fontWeight: 600, whiteSpace: `nowrap`, overflow: `hidden`, textOverflow: `ellipsis` } }, s.label),
          elQ(`div`, { key: `v`, style: { fontSize: 60, fontWeight: 800, lineHeight: 1.1, fontVariantNumeric: `tabular-nums`, color: callCol(s.call) } }, `$${Y(s.tgt, 1)}`),
          elQ(`div`, { key: `s`, style: { fontSize: 18, color: `var(--color-muted)`, marginTop: 4 } }, `${s.call} · 較現價 ${s.tgt >= P ? `+` : `−`}${hA(Math.abs(s.tgt / P - 1) * 100, 0)}`)
        ]))
      ]),
      // 價格刻度：三情境目標價、賣出門檻、現價
      elQ(`div`, { key: `r`, className: `sumQ-a sumQ-a3`, style: { position: `relative`, height: 100, margin: `34px 40px 0` } }, [
        elQ(`div`, { key: `ax`, style: { position: `absolute`, left: 0, right: 0, top: 32, height: 10, borderRadius: 5, background: `linear-gradient(90deg, var(--color-bad) 0%, var(--color-bad) ${X(TR.th)}, var(--color-surface) ${X(TR.th)})`, opacity: .35 } }),
        elQ(`div`, { key: `t`, style: { position: `absolute`, left: X(TR.th), top: 18, height: 38, width: 2, background: `var(--color-bad)` } }),
        elQ(`div`, { key: `tl`, style: { position: `absolute`, left: X(TR.th), top: 60, transform: `translateX(-50%)`, fontSize: 18, color: `var(--color-bad)`, whiteSpace: `nowrap` } }, `賣出門檻 $${Y(TR.th, 1)}`),
        // 標籤相近時（距離 < 刻度 7%）改放在軸下方，避免重疊
        ...scen.map((s, i) => { let lv = (() => { let L = []; scen.forEach((q, j) => { let c = [0, 1, 2].find(l => !scen.some((w, k) => k < j && L[k] === l && Math.abs(w.tgt - q.tgt) < hi * .07)); L.push(c ?? 2) }); return L })()[i], low = lv > 0; // v0.2a：三個標籤相近時分三層（上、下、下二），避免重疊
          return elQ(`div`, { key: s.sc, style: { position: `absolute`, left: X(s.tgt), top: low ? 28 : 0, transform: `translateX(-50%)`, textAlign: `center`, display: `flex`, flexDirection: low ? `column-reverse` : `column`, alignItems: `center` } }, [
            elQ(`div`, { key: `l`, style: { fontSize: 18, lineHeight: `22px`, fontWeight: 700, whiteSpace: `nowrap`, color: s.sc === curSc ? `var(--color-accent)` : `var(--color-fg)`, marginTop: lv === 2 ? 28 : low ? 4 : 0 } }, `${s.label.split(/\s/)[0]} $${Y(s.tgt, 1)}`),
            elQ(`div`, { key: `d`, style: { width: 18, height: 18, borderRadius: 9, background: callCol(s.call), marginTop: low ? 0 : 6, border: `3px solid var(--color-card)` } })
          ]); }),
        elQ(`div`, { key: `p`, style: { position: `absolute`, left: X(P), top: 0, transform: `translateX(-50%)`, textAlign: `center` } }, [
          elQ(`div`, { key: `l`, style: { fontSize: 18, lineHeight: `22px`, fontWeight: 700, whiteSpace: `nowrap` } }, `現價 $${Y(P, 1)}`),
          elQ(`div`, { key: `d`, style: { width: 0, height: 0, margin: `6px auto 0`, borderLeft: `11px solid transparent`, borderRight: `11px solid transparent`, borderTop: `18px solid var(--color-ink)` } })
        ])
      ]),
      elQ(NoteQ, { key: `n` }, `三個擴張情境的目標價${scen.every(s => s.tgt < TR.th) ? `全部低於賣出門檻` : `並非全部低於賣出門檻`}：結論「${f.call.call}」（點位＝${scLabel}）。`)
    ], 56);
  }

  // 2｜Nebius 怎麼賺錢：已連網 MW × 每 MW 年收入
  {
    let acc = d.m.accepted, bil = y.map(t => t.avgBillable), rv4 = d.m.revMW[4] * (e.revScale ?? 1) * 1e3, R = f.fwd.map(x => x.fyRevenue), EB = f.fwd.map(x => x.fyEbitda);
    let mxA = Math.max(...acc, 1), mxR = Math.max(...R, .1);
    let box = (k, icon, label, value, sub, col) => elQ(`div`, { key: k, className: `sumQ-a sumQ-a2`, style: { flex: 1, minWidth: 0, padding: `14px 18px`, borderRadius: 12, background: `var(--color-surface)`, borderTop: `5px solid ${col}` } }, [
      elQ(`div`, { key: `l`, style: { display: `flex`, alignItems: `center`, gap: 8, fontSize: 19, fontWeight: 600, color: `var(--color-muted)` } }, [elQ(IconQ, { key: `i`, n: icon, size: 24, color: col }), label]),
      elQ(`div`, { key: `v`, style: { fontSize: 56, fontWeight: 800, lineHeight: 1.1, marginTop: 4, fontVariantNumeric: `tabular-nums`, whiteSpace: `nowrap` } }, value),
      elQ(`div`, { key: `s`, style: { fontSize: 18, color: `var(--color-muted)`, marginTop: 2 } }, sub)
    ]), op = (k, t) => elQ(`div`, { key: k, style: { fontSize: 44, fontWeight: 300, color: `var(--color-subtle)`, alignSelf: `center` } }, t);
    M(`2｜${COMPANY_DATA.meta.company} 怎麼賺錢`, `營收＝已連網 MW × 每 MW 年收入：${PERIODS[4]} 約 ${Y(bil[4], 0)} MW 計費 × $${Y(rv4, 1)}m ＝ ${money(R[4])}`, [
      elQ(`div`, { key: `c`, style: { display: `flex`, gap: 14, alignItems: `stretch` } }, [
        box(`a`, `bolt`, `${PERIODS[4]} 年底已連網`, `${Y(acc[4], 0)} MW`, `平均計費 ${Y(bil[4], 0)} MW（在役比例）`, COLQ.debt), op(`x`, `×`),
        box(`b`, `coin`, `每 MW 年收入`, `$${Y(rv4, 1)}m`, PMW_REVQ === `tkAnchor` ? (A => `Tokenomics 錨 $${Y(A.anchor[4] * 1e3, 1)}m × k ${Y(A.k[4], 2)}`)(tkAnchorQ(curSc || `base`)) : `Tokenomics 正向推導（非公司 ACV）`, COLQ.op), op(`=`, `=`),
        box(`c`, `chart`, `${PERIODS[4]} 營收`, money(R[4]), `EBITDA ${money(EB[4])}（${hA(y[4].ebM * 100, 0)}）`, COLQ.prepay)
      ]),
      elQ(`div`, { key: `g`, className: `sumQ-a sumQ-a3`, style: { display: `grid`, gridTemplateColumns: `repeat(5, 1fr)`, gap: 26, alignItems: `end`, height: 236, marginTop: 22, padding: `0 30px` } }, y.map((t, i) => elQ(`div`, { key: i, style: { display: `flex`, flexDirection: `column`, alignItems: `center`, justifyContent: `flex-end`, height: `100%` } }, [
        elQ(`div`, { key: `v`, style: { fontSize: 20, fontWeight: 700, fontVariantNumeric: `tabular-nums` } }, money(R[i])),
        elQ(`div`, { key: `b`, style: { display: `flex`, gap: 6, alignItems: `flex-end`, height: 150, marginTop: 4 } }, [
          elQ(`div`, { key: `m`, className: `sumQ-gy`, style: { width: 40, height: acc[i] / mxA * 150, background: COLQ.debt, opacity: .35, borderRadius: `4px 4px 0 0` } }),
          elQ(`div`, { key: `r`, className: `sumQ-gy`, style: { width: 40, height: R[i] / mxR * 150, background: COLQ.prepay, borderRadius: `4px 4px 0 0` } })
        ]),
        elQ(`div`, { key: `y`, style: { fontSize: 19, fontWeight: 600, marginTop: 6 } }, `${PERIODS[i]} · ${Y(acc[i], 0)} MW`)
      ]))),
      elQ(`div`, { key: `l`, className: `sumQ-a sumQ-a3`, style: { marginTop: 10, paddingLeft: 30 } }, elQ(LegQ, { items: [[`已連網 MW（淺）`, `${COLQ.debt}66`], [`營收 US$bn`, COLQ.prepay]] })),
      elQ(NoteQ, { key: `n` }, `RPO 只作對照，不驅動營收；可計費 MW 依在役比例與爬坡係數逐年追上已連網 MW。`)
    ]);
  }

  // 3｜預付款機制：預付 vs 預付設 0（同一引擎，只把客戶預付比率設為 0）
  {
    let pairs = [
      [`新股（五期）`, `百萬股`, T.newShares * 1e3, p0.d.totals.newShares * 1e3, x => `${Y(x, 0)}m`],
      [`融資前缺口`, `US$bn`, gap0, Math.max(0, -p0.d.totals.preFinEnd), x => money(x)],
      [`目標價（每股）`, `US$`, tgt, p0.tgt, x => `$${Y(x, 1)}`]
    ];
    let dSh = (p0.d.totals.newShares - T.newShares) * 1e3, dT = tgt - p0.tgt;
    let step = (k, icon, a, b, col) => elQ(`div`, { key: k, className: `sumQ-a sumQ-a2`, style: { display: `flex`, alignItems: `center`, gap: 12, padding: `10px 14px`, borderRadius: 10, background: `var(--color-surface)` } }, [
      elQ(`div`, { key: `i`, style: { width: 44, height: 44, borderRadius: 22, background: `${col}22`, display: `grid`, placeItems: `center` } }, elQ(IconQ, { n: icon, size: 26, color: col })),
      elQ(`div`, { key: `t` }, [elQ(`div`, { key: `a`, style: { fontSize: 20, fontWeight: 700 } }, a), elQ(`div`, { key: `b`, style: { fontSize: 18, color: `var(--color-muted)` } }, b)])
    ]);
    let arr = k => elQ(`div`, { key: k, style: { textAlign: `center`, fontSize: 22, color: `var(--color-subtle)`, lineHeight: 1 } }, `↓`);
    M(`3｜預付款機制`, `客戶先付錢：少發 ${Y(dSh, 0)}m 股、目標價高 $${Y(dT, 1)}——但缺口仍有 ${money(gap0)}`, [
      elQ(`div`, { key: `g`, style: { display: `grid`, gridTemplateColumns: `400px 1fr`, gap: 36, flex: 1, minHeight: 0 } }, [
        elQ(`div`, { key: `f`, style: { display: `flex`, flexDirection: `column`, gap: 6 } }, [
          step(`1`, `users`, `客戶簽約`, `輸入：${hA((COMPANY_DATA.defaults.prepay?.shareOfDeals ?? 0) * 100, 0)} 合約含預付 × 覆蓋 ${hA((COMPANY_DATA.defaults.prepay?.capexCover ?? 0) * 100, 0)}`, COLQ.op), arr(`a1`),
          step(`2`, `coin`, `建置前先收現`, `五期約占毛 CapEx ${hA(cover * 100, 0)}`, COLQ.prepay), arr(`a2`),
          step(`3`, `server`, `用預付款蓋 GPU`, `五期預付流入 ${money(preIn)}`, COLQ.debt), arr(`a3`),
          step(`4`, `doc`, `之後認列為營收`, `${e.prepay?.recogYears ?? `—`} 年攤入（非現金）`, COLQ.cash)
        ]),
        elQ(`div`, { key: `b`, style: { display: `flex`, flexDirection: `column`, gap: 16 } }, [
          elQ(LegQ, { key: `l`, items: [[`有預付（模型）`, COLQ.prepay], [`預付設為 0（對照）`, COLQ.off]] }),
          ...pairs.map(([lab, u, a, b, fm], i) => {
            let mx = Math.max(a, b, 1e-9);
            return elQ(`div`, { key: lab, className: `sumQ-a sumQ-a3` }, [
              elQ(`div`, { key: `l`, style: { fontSize: 20, fontWeight: 700, marginBottom: 4 } }, lab),
              ...[[a, COLQ.prepay], [b, COLQ.off]].map(([x, c], j) => elQ(`div`, { key: j, style: { display: `grid`, gridTemplateColumns: `1fr 120px`, alignItems: `center`, gap: 12, marginTop: 4 } }, [
                elQ(`div`, { key: `b`, style: { height: 26, background: `var(--color-surface)`, borderRadius: 4 } }, elQ(`div`, { className: `sumQ-g`, style: { width: `${x / mx * 100}%`, height: `100%`, background: c, borderRadius: 4 } })),
                elQ(`div`, { key: `v`, style: { fontSize: 22, fontWeight: 700, textAlign: `right`, fontVariantNumeric: `tabular-nums` } }, fm(x))
              ]))
            ]);
          })
        ])
      ]),
      elQ(NoteQ, { key: `n` }, `對照組只把客戶預付比率設為 0，其餘輸入不變（同一引擎試算）。`)
    ]);
  }

  // 4｜錢從哪裡來：五期資金來源堆疊（預付 → 營運現金 → 現金 → 資產擔保債 → 可轉債 → 新股 → 高息債）
  {
    let U = sumY(`uses`), opN = sumY(`sourcesOp`) - preIn, cashN = Math.max(0, T.atm + Math.max(0, -sumY(`gap`)));
    let parts = [[`客戶預付`, preIn, COLQ.prepay, `預付`], [`營運現金`, Math.max(0, opN), COLQ.op, `營運`], [`現金與期後募資`, cashN, COLQ.cash, `現金`], [`資產擔保債`, T.newDebt, COLQ.debt, `擔保債`],
      [`可轉債`, T.convNew, COLQ.conv, `可轉債`], [`新股`, T.equity, COLQ.eq, `新股`], [`高息債`, T.junk, COLQ.junk, `高息`]];
    let tot = parts.reduce((a, x) => a + x[1], 0) || 1, ext = T.convNew + T.equity;
    let useP = [[`毛 CapEx`, T.gross, COLQ.need], [`租金、利息、還本等`, Math.max(0, U - T.gross), COLQ.off]];
    let bar = (k, ps, base, h, lab, named) => elQ(`div`, { key: k, className: `sumQ-a sumQ-a2` }, [
      elQ(`div`, { key: `t`, style: { fontSize: 20, fontWeight: 700, marginBottom: 6 } }, lab),
      // v0.2 第 2 輪：色塊內同時標類別（短名）與金額，避免兩段同為「12」時無法辨識
      elQ(`div`, { key: `b`, className: `sumQ-g`, style: { display: `flex`, height: h, borderRadius: 8, overflow: `hidden`, width: `100%` } }, ps.filter(x => x[1] > .05).map(([a, x, c, sh]) =>
        elQ(`div`, { key: a, title: a, style: { width: `${x / base * 100}%`, background: c, color: `#fff`, display: `flex`, flexDirection: named ? `row` : `column`, gap: named ? 8 : 0, alignItems: `center`, justifyContent: `center`, fontSize: named ? 20 : 18, lineHeight: 1.15, fontWeight: 700, overflow: `hidden`, whiteSpace: `nowrap`, borderRight: `2px solid var(--color-card)` } },
          x / base > .06 ? [named ? (x / base > .25 ? elQ(`span`, { key: `n` }, a) : null) : elQ(`span`, { key: `n`, style: { fontWeight: 600, opacity: .92 } }, sh || a), elQ(`span`, { key: `v` }, Y(x, 1))] : ``)))
    ]);
    M(`4｜錢從哪裡來`, `五期要花 ${money(U, 0)}：預付付 ${money(preIn, 0)}，可轉債與新股還要補 ${money(ext, 0)}`, [
      bar(`s`, parts, tot, 74, `資金來源（${PERIODS[0]}–${PERIODS[4]} 合計，US$bn；依瀑布順序由左至右）`),
      elQ(`div`, { key: `l`, className: `sumQ-a sumQ-a2`, style: { marginTop: 10 } }, elQ(LegQ, { items: parts.filter(x => x[1] > .05).map(([a, x, c]) => [`${a} ${Y(x, 1)}`, c]) })),
      elQ(`div`, { key: `u`, style: { marginTop: 18 } }, bar(`u`, useP, tot, 40, `資金用途（同期合計；淺色＝租金、利息、還本等）`, !0)),
      elQ(`div`, { key: `k`, style: { display: `grid`, gridTemplateColumns: `repeat(3, 1fr)`, gap: 24, marginTop: 22 } }, [
        elQ(BigQ, { key: 1, label: `客戶預付`, value: money(preIn), color: COLQ.prepay, icon: `coin`, sub: `覆蓋毛 CapEx ${hA(cover * 100, 0)}` }),
        elQ(BigQ, { key: 2, label: `可轉債（新發）`, value: money(T.convNew), color: COLQ.conv, icon: `doc`, sub: `每年上限 ${money(e.cvCap ?? 0)}` }),
        elQ(BigQ, { key: 3, label: `新股`, value: `${shM(T.newShares)} 股`, color: COLQ.eq, icon: `users`, sub: `現有股數的 ${hA(T.newShares / o.shares * 100, 0)}（${money(T.equity)}）` })
      ]),
      elQ(NoteQ, { key: `n` }, `每期在需要前先融足，期末現金不低於 ${money(e.minCash ?? 2)}；新融資的利息也由瀑布支應。`)
    ]);
  }

  // 5｜關鍵疑點：實現單價 vs 正向推導 vs 公司 ACV vs 現價所需
  {
    let K = PRICE_CHK_Q, rv4 = d.m.revMW[4] * (e.revScale ?? 1) * 1e3, need = rv && Number.isFinite(rv.Rt) ? rv4 * rv.Rt : NaN,
      TK5 = PMW_REVQ === `tkAnchor` ? tkAnchorQ(curSc || `base`) : null, // v0.2a：刻度尺＝Q2 實現／Tokenomics 錨／錨 × k（模型）／公司 ACV／現價所需
      der = TK5 ? [[`Tokenomics 錨（打平租金）`, TK5.anchor[4] * 1e3, `anc`, COLQ.cash, !1, !0], [`錨 × k ${Y(rv4 / (TK5.anchor[4] * 1e3), 2)}（模型）`, rv4, `mod`, COLQ.model, !0, !1]]
        : [`low`, `base`, `high`].map(sc => [`推導・${SCENARIOS[sc].label.split(/\s/)[0]}`, SCENARIOS[sc].rev[4] * 1e3, sc, sc === curSc ? COLQ.model : `${COLQ.model}99`, sc === curSc]);
    let pts = [...der.map(x => x[1]), K ? K.realized : 0, ...(K?.acv || []), Number.isFinite(need) ? need : 0];
    let hi = Math.ceil(Math.max(...pts) * 1.12 / 5) * 5, X = x => `${x / hi * 100}%`;
    let ticks = []; for (let t = 0; t <= hi; t += 5) ticks.push(t);
    let mark = (k, x, lab, val, col, up, shape, big) => elQ(`div`, { key: k, style: { position: `absolute`, left: X(x), top: 0, bottom: 0, width: 0 } }, [
      elQ(`div`, { key: `d`, style: { position: `absolute`, top: 132, left: big ? -14 : -11, width: big ? 28 : 22, height: big ? 28 : 22, borderRadius: shape === `d` ? 3 : 14, transform: shape === `d` ? `rotate(45deg)` : `none`, background: col, border: `3px solid var(--color-card)`, boxShadow: `0 0 0 1px ${col}` } }),
      elQ(`div`, { key: `t`, style: { position: `absolute`, top: up ? 62 : 172, left: 0, transform: `translateX(-50%)`, textAlign: `center`, whiteSpace: `nowrap` } }, [
        elQ(`div`, { key: `v`, style: { fontSize: big ? 30 : 24, fontWeight: 800, color: col, fontVariantNumeric: `tabular-nums` } }, val),
        elQ(`div`, { key: `l`, style: { fontSize: 18, color: `var(--color-muted)` } }, lab)
      ])
    ]);
    M(`5｜關鍵疑點`, TK5 && K ? `實際 $${Y(K.realized, 1)}m、打平租金 $${Y(TK5.anchor[4] * 1e3, 1)}m、模型 $${Y(rv4, 1)}m${Number.isFinite(need) ? `、現價要 $${Y(need, 1)}m` : ``}` : K ? `每 MW 實際只收 $${Y(K.realized, 1)}m，模型用 $${Y(rv4, 1)}m${Number.isFinite(need) ? `，現價要 $${Y(need, 1)}m` : ``}` : `每 MW 年收入：模型 $${Y(rv4, 1)}m`, [
      elQ(`div`, { key: `r`, className: `sumQ-a sumQ-a2`, style: { position: `relative`, height: 290, margin: `0 40px` } }, [
        K?.acv ? elQ(`div`, { key: `acv`, style: { position: `absolute`, left: X(K.acv[0]), width: `calc(${X(K.acv[1])} - ${X(K.acv[0])})`, top: 110, height: 33, background: `${COLQ.cons}22`, border: `2px dashed ${COLQ.cons}`, borderRadius: 8 } }) : null, // v0.2 第 2 輪：框只在軸上方，不蓋住刻度數字
        K?.acv ? elQ(`div`, { key: `acvl`, style: { position: `absolute`, left: `calc((${X(K.acv[0])} + ${X(K.acv[1])}) / 2)`, transform: `translateX(-50%)`, top: 0, textAlign: `center`, whiteSpace: `nowrap` } }, [
          elQ(`div`, { key: `v`, style: { fontSize: 24, fontWeight: 800, color: COLQ.cons } }, `$${Y(K.acv[0], 0)}–${Y(K.acv[1], 0)}m`),
          elQ(`div`, { key: `l`, style: { fontSize: 18, color: `var(--color-muted)` } }, `公司新約 ACV（自述）`)]) : null,
        elQ(`div`, { key: `ax`, style: { position: `absolute`, left: 0, right: 0, top: 141, height: 4, background: `var(--color-border)` } }),
        ...ticks.map(t => elQ(`div`, { key: `t${t}`, style: { position: `absolute`, left: X(t), top: 150, transform: `translateX(-50%)`, fontSize: 18, color: `var(--color-subtle)` } }, `${t}`)),
        K ? mark(`real`, K.realized, `實現（最新一季）`, `$${Y(K.realized, 1)}m`, COLQ.real, !0, `c`, !0) : null,
        ...der.map(([lab, x, sc, col, big, up]) => mark(sc, x, lab, `$${Y(x, 1)}m`, col, !!up, `c`, big)), // v0.2a 第 2 輪：錨標籤放軸上方（模型低於錨時兩標籤相鄰）
        Number.isFinite(need) ? mark(`need`, need, `現價所需`, `$${Y(need, 1)}m`, COLQ.need, !0, `d`, !0) : null
      ]),
      elQ(`div`, { key: `c`, className: `sumQ-a sumQ-a3`, style: { display: `flex`, gap: 14, marginTop: 4, flexWrap: `wrap` } }, (TK5 ? [[`爬坡時點`, K ? `Q2 差距 ${hA((LATEST_Q.revenue * 4e3 / e.billableOpen - K.realized) / (TK5.rev[0] * 1e3 - K.realized) * 100, 0)} 來自計費 MW 少於在役` : `營收落後於交付`], [`定價倍數 k`, `新增產能按長約價：k＝${Y(TK5.k[4], 2)}（隨需 ${hA(TK5.od * 100, 0)}）`], [`MW 口徑`, `在役 MW 為內插估計`]]
        : [[`爬坡時點`, `模型採用：營收落後於交付`], [`舊世代機隊`, `H100／H200 單價較低`], [`MW 口徑`, `在役 MW 為內插估計`]]).map(([a, b], i) =>
        elQ(`div`, { key: a, style: { flex: 1, padding: `10px 14px`, borderRadius: 10, background: `var(--color-surface)`, borderLeft: `4px solid ${i ? COLQ.cash : COLQ.model}` } }, [
          elQ(`div`, { key: `a`, style: { fontSize: 20, fontWeight: 700 } }, a), elQ(`div`, { key: `b`, style: { fontSize: 18, color: `var(--color-muted)` } }, b)]))),
      elQ(NoteQ, { key: `n` }, TK5 && K ? `US$m／MW·年；模型、錨與現價所需為 ${PERIODS[4]} 值。實現＝最新一季營收 × 4 ÷ 在役 MW 估計 ${Y(K.inServiceMw, 0)}；錨＝Tokenomics IF_HoldEcon（WACC 10% 打平租金）在役世代加權。`
        : K ? `US$m／MW·年。實現＝最新一季營收 × 4 ÷ 在役 MW 估計 ${Y(K.inServiceMw, 0)}。` : `US$m／MW·年。`)
    ]);
  }

  // 6｜現價隱含什麼：反向 DCF
  {
    let ok = rv && Number.isFinite(rv.Rt), rv4 = rv ? rv.rev30 : d.m.revMW[4] * 1e3, lo = rv ? rv.ebs[0] : NaN, hi2 = rv ? rv.ebs[2] : NaN;
    let G = x => `${Math.max(0, Math.min(100, x * 100))}%`;
    M(`6｜現價隱含什麼`, ok ? `現價 $${Y(P, 2)} 要成立：每 MW 年收入得是模型的 ${Y(rv.Rt, 1)} 倍${Number.isFinite(rv.Eb) ? `，或 EBITDA 率 ${hA(rv.Eb * 100, 0)}` : ``}` : `現價隱含什麼（計算中）`, ok ? [
      elQ(`div`, { key: `g`, style: { display: `grid`, gridTemplateColumns: `1fr 1fr`, gap: 48, marginTop: 4 } }, [
        elQ(`div`, { key: `a` }, [
          elQ(BigQ, { key: `b`, label: `每 MW 年收入（${PERIODS[4]}）`, icon: `coin`, value: `×${Y(rv.Rt, 1)}`, color: COLQ.eq, sub: `$${Y(rv4, 1)}m → $${Y(rv4 * rv.Rt, 1)}m（加權目標價＝現價）` }),
          ...[[`模型`, rv4, COLQ.model], [`現價所需`, rv4 * rv.Rt, COLQ.eq]].map(([a, x, c]) => elQ(`div`, { key: a, className: `sumQ-a sumQ-a3`, style: { display: `grid`, gridTemplateColumns: `110px 1fr`, gap: 12, alignItems: `center`, marginTop: 14 } }, [
            elQ(`div`, { key: `l`, style: { fontSize: 19, fontWeight: 600 } }, a),
            elQ(`div`, { key: `b`, style: { height: 30, background: `var(--color-surface)`, borderRadius: 4 } }, elQ(`div`, { className: `sumQ-g`, style: { width: `${x / (rv4 * Math.max(rv.Rt, 1)) * 100}%`, height: `100%`, background: c, borderRadius: 4 } }))
          ]))
        ]),
        elQ(`div`, { key: `b` }, [
          elQ(BigQ, { key: `b`, label: `或穩態 EBITDA 率`, icon: `target`, value: Number.isFinite(rv.Eb) ? hA(rv.Eb * 100, 0) : `無解`, color: COLQ.eq, sub: `模型 ${hA(rv.eb30 * 100, 0)}；DCF＝現價、其他不變` }),
          elQ(`div`, { key: `gg`, className: `sumQ-a sumQ-a3`, style: { position: `relative`, height: 64, marginTop: 40 } }, [
            elQ(`div`, { key: `ax`, style: { position: `absolute`, left: 0, right: 0, top: 10, height: 26, background: `var(--color-surface)`, borderRadius: 13 } }),
            elQ(`div`, { key: `bd`, style: { position: `absolute`, left: G(lo), width: `calc(${G(hi2)} - ${G(lo)})`, top: 10, height: 26, background: `${COLQ.model}55`, borderRadius: 4 } }),
            elQ(`div`, { key: `bl`, style: { position: `absolute`, left: G(lo), top: 40, fontSize: 18, color: `var(--color-muted)`, whiteSpace: `nowrap` } }, `可觀察 neocloud ${hA(lo * 100, 0)}–${hA(hi2 * 100, 0)}`),
            Number.isFinite(rv.Eb) ? elQ(`div`, { key: `nd`, style: { position: `absolute`, left: G(rv.Eb), top: 4, width: 4, height: 38, marginLeft: -2, background: COLQ.eq } }) : null,
            Number.isFinite(rv.Eb) ? elQ(`div`, { key: `ndl`, style: { position: `absolute`, left: G(rv.Eb), top: -22, transform: `translateX(-50%)`, fontSize: 18, fontWeight: 700, color: COLQ.eq, whiteSpace: `nowrap` } }, `所需`) : null,
            elQ(`div`, { key: `md`, style: { position: `absolute`, left: G(rv.eb30), top: 4, width: 4, height: 38, marginLeft: -2, background: COLQ.model } }),
            elQ(`div`, { key: `mdl`, style: { position: `absolute`, left: G(rv.eb30), top: -22, transform: `translateX(-50%)`, fontSize: 18, fontWeight: 700, color: COLQ.model, whiteSpace: `nowrap` } }, `模型`)
          ])
        ])
      ]),
      elQ(`div`, { key: `c`, className: `sumQ-a sumQ-a3`, style: { marginTop: 26, padding: `14px 20px`, borderRadius: 12, background: `var(--color-surface)`, fontSize: 22, fontWeight: 600, lineHeight: 1.5 } },
        Number.isFinite(rv.Eb) && rv.Eb > hi2 ? `現價要求的利潤率超出所有可觀察 neocloud；只靠營運效率不夠，得靠更高的每 MW 單價。` : `現價所需條件落在可觀察範圍內，需逐項檢驗。`),
      elQ(NoteQ, { key: `n` }, `DCF 單獨＝現價需每 MW 年收入 ×${Y(rv.R, 2)}；矩陣與解法見附錄「反向 DCF」。`)
    ] : elQ(`p`, { style: { fontSize: 20 } }, `計算中…`));
  }

  // 7｜容量 × 價格（v0.2a）：3 × 3 目標價；容量情境只改 MW，價格情境只改定價倍數 k
  if (km) {
    let K3 = [`low`, `base`, `high`], AM = COMPANY_DATA.pricing.anchorMultiple, pxN = { low: `價格低`, base: `價格基準`, high: `價格高` }, bs = curSc || `base`,
      lo = km[bs].low.tgt, hi3 = km[bs].high.tgt, cLo = Math.min(...K3.map(s => km[s].base.tgt)), cHi = Math.max(...K3.map(s => km[s].base.tgt));
    let cell = (sc, px) => { let x = km[sc][px], c = x.tgt < TR.th ? COLQ.real : x.tgt > P ? `var(--color-ok)` : COLQ.conv, cur = sc === curSc && px === `base`;
      return elQ(`div`, { key: sc + px, className: `sumQ-a sumQ-a3`, style: { padding: `10px 16px`, borderRadius: 12, background: `var(--color-surface)`, border: cur ? `3px solid var(--color-accent)` : `3px solid transparent`, textAlign: `center` } }, [
        elQ(`div`, { key: `v`, style: { fontSize: 44, fontWeight: 800, color: c, fontVariantNumeric: `tabular-nums`, lineHeight: 1.1 } }, `$${Y(x.tgt, 1)}`),
        elQ(`div`, { key: `s`, style: { fontSize: 18, color: `var(--color-muted)` } }, `新股 $${Y(x.eq, 1)}bn`)]); };
    M(`7｜容量 × 價格`, `價格比容量重要：同一容量下 k 由低到高，目標價 $${Y(lo, 0)} → $${Y(hi3, 0)}`, [
      elQ(`div`, { key: `g`, style: { display: `grid`, gridTemplateColumns: `260px 1fr 1fr 1fr`, gap: 14, alignItems: `center` } }, [
        elQ(`div`, { key: `h0` }),
        ...K3.map(px => elQ(`div`, { key: `h` + px, className: `sumQ-a sumQ-a2`, style: { textAlign: `center` } }, [
          elQ(`div`, { key: `a`, style: { fontSize: 22, fontWeight: 700 } }, pxN[px]),
          elQ(`div`, { key: `b`, style: { fontSize: 18, color: `var(--color-muted)` } }, `k_長約 ${Y(AM.long[px], 2)}`)])),
        ...K3.flatMap(sc => [elQ(`div`, { key: `r` + sc, className: `sumQ-a sumQ-a2`, style: { fontSize: 20, fontWeight: 700 } }, SCENARIOS[sc].label), ...K3.map(px => cell(sc, px))])
      ]),
      elQ(`div`, { key: `l`, className: `sumQ-a sumQ-a3`, style: { marginTop: 14 } }, elQ(LegQ, { items: [[`低於賣出門檻 $${Y(TR.th, 1)}`, COLQ.real], [`門檻與現價之間`, COLQ.conv], [`高於現價 $${Y(P, 1)}`, `var(--color-ok)`]] })),
      elQ(NoteQ, { key: `n` }, `每 MW 年收入＝Tokenomics 錨 × k；三個容量情境的價格基準欄目標價 $${Y(cLo, 1)}–$${Y(cHi, 1)}。粗框＝目前情境。新股＝五期股權募資。`)
    ]);
  }

  // 7｜與市場共識的差距：倍數、利潤率、稀釋
  {
    let last = cv.rows[cv.rows.length - 1], PT = CONSENSUS.priceTarget;
    let row = (k, lab, a, b, fm, extra) => {
      let mx = Math.max(a ?? 0, b ?? 0, 1e-9);
      return elQ(`div`, { key: k, className: `sumQ-a sumQ-a3`, style: { display: `grid`, gridTemplateColumns: `300px 1fr`, gap: 20, alignItems: `center`, padding: `10px 0`, borderTop: `1px solid var(--color-border)` } }, [
        elQ(`div`, { key: `l` }, [elQ(`div`, { key: `a`, style: { fontSize: 22, fontWeight: 700 } }, lab), extra ? elQ(`div`, { key: `b`, style: { fontSize: 18, color: `var(--color-muted)` } }, extra) : null]),
        elQ(`div`, { key: `b` }, [[`模型`, a, COLQ.model], [`共識`, b, COLQ.cons]].map(([n, x, c]) => elQ(`div`, { key: n, style: { display: `grid`, gridTemplateColumns: `54px 1fr 120px`, gap: 10, alignItems: `center`, marginTop: 3 } }, [
          elQ(`div`, { key: `n`, style: { fontSize: 18, color: `var(--color-muted)` } }, n),
          elQ(`div`, { key: `b`, style: { height: 24, background: `var(--color-surface)`, borderRadius: 4 } }, x == null ? null : elQ(`div`, { className: `sumQ-g`, style: { width: `${x / mx * 100}%`, height: `100%`, background: c, borderRadius: 4 } })),
          elQ(`div`, { key: `v`, style: { fontSize: 22, fontWeight: 700, textAlign: `right`, fontVariantNumeric: `tabular-nums`, color: x == null ? `var(--color-subtle)` : `var(--color-fg)` } }, x == null ? `未揭露` : fm(x))
        ])))
      ]);
    };
    M(`${km ? 8 : 7}｜與市場共識的差距`, `模型 $${Y(TR.pt, 1)} vs 共識 $${Y(PT.mean, 1)}：差在倍數、利潤率、稀釋`, [
      elQ(`div`, { key: `t`, className: `sumQ-a sumQ-a2`, style: { display: `flex`, alignItems: `center`, gap: 28, marginBottom: 10 } }, [
        elQ(BigQ, { key: `a`, label: `模型點位`, value: `$${Y(TR.pt, 1)}`, color: COLQ.model }),
        elQ(`div`, { key: `x`, style: { fontSize: 40, color: `var(--color-subtle)` } }, `→`),
        elQ(BigQ, { key: `b`, label: `共識平均目標價（${PT.analysts} 家）`, value: `$${Y(PT.mean, 1)}`, color: COLQ.cons }),
        elQ(`div`, { key: `g`, style: { marginLeft: `auto`, fontSize: 22, fontWeight: 700, color: `var(--color-bad)`, textAlign: `right` } }, `模型低 ${hA(Math.abs(TR.pt / PT.mean - 1) * 100, 0)}`)
      ]),
      row(`m`, `${last.yr} EV/EBITDA 倍數`, o.evEbitda, cv.impTgt, x => `${Y(x, 1)}x`, `現價隱含 ${Y(cv.impPx, 1)}x`),
      row(`e`, `${last.yr} EBITDA 率`, last.m.ebM, last.c.ebM, x => hA(x * 100, 1), `模型取可觀察 neocloud 中點`),
      row(`s`, `新股（五期）`, T.newShares * 1e3, null, x => `${Y(x, 0)}m`, `占現有股數 ${hA(T.newShares / o.shares * 100, 0)}`),
      elQ(NoteQ, { key: `n` }, `另：DCF 腿${f.d.perShareRaw < 0 ? `為負、以 0 截斷` : `為正`}，仍占 ${hA((f.call.weights?.dcf ?? .45) * 100, 0)} 權重。`)
    ]);
  }

  // 8｜驗證點：下一季財報要看的數字
  {
    let keys = [...new Set([`mw`, ...((COMPANY_DATA.quarterly || {}).keyMetrics || [])])];
    let rows = qv ? keys.map(k => qv.rows.find(m => m.key === k)).filter(Boolean).slice(0, 4) : [];
    let why = { mw: `→ 實現單價是爬坡還是低價`, revenue: `→ 是否落入全年指引`, adjEbitda: `→ 利潤率是否守住`, capex: `→ 建置速度與融資需求`, ebitdaMargin: `→ 利潤率是否守住` };
    let fv = (m, x) => x == null ? `—` : m.unit === `MW` ? `${Y(x, 0)} MW` : m.unit === `%` ? hA(x * 100, 1) : money(x, 2);
    M(`${km ? 9 : 8}｜驗證點`, qv ? `下一個驗證點：${qv.focus.label} 財報${qv.focus.reportNote ? `（${qv.focus.reportNote}）` : ``}，看這 ${rows.length} 個數字` : `下一個驗證點：${CALL_FACTS.nextEarn || `下一季財報`}`, [
      elQ(`div`, { key: `g`, style: { display: `grid`, gridTemplateColumns: `1fr 1fr`, gridAutoRows: `1fr`, gap: 18, flex: 1, minHeight: 0, maxHeight: 380 } }, rows.map((m, i) => {
        let r = m.q[qv.fi];
        return elQ(`div`, { key: m.key, className: `sumQ-a sumQ-a${2 + (i % 3)}`, style: { padding: `14px 20px`, borderRadius: 12, background: `var(--color-surface)`, borderLeft: `6px solid ${[COLQ.real, COLQ.model, COLQ.prepay, COLQ.conv][i]}`, display: `flex`, flexDirection: `column`, justifyContent: `center`, minHeight: 0 } }, [
          elQ(`div`, { key: `l`, style: { display: `flex`, alignItems: `center`, gap: 8, fontSize: 21, fontWeight: 700 } }, [elQ(IconQ, { key: `i`, n: [`eye`, `chart`, `target`, `server`][i], size: 24, color: [COLQ.real, COLQ.model, COLQ.prepay, COLQ.conv][i] }), m.label]),
          elQ(`div`, { key: `v`, style: { display: `flex`, alignItems: `baseline`, gap: 14, marginTop: 2 } }, [
            elQ(`span`, { key: `a`, style: { fontSize: 56, fontWeight: 800, fontVariantNumeric: `tabular-nums`, lineHeight: 1.1 } }, fv(m, r.model)),
            elQ(`span`, { key: `b`, style: { fontSize: 19, color: `var(--color-muted)` } }, `模型${r.cons != null ? `｜共識 ${fv(m, r.cons)}` : ``}${r.guide ? `｜指引 ${rngTxtQ(r.guide, m.unit)}` : ``}`)
          ]),
          elQ(`div`, { key: `w`, style: { fontSize: 18, color: `var(--color-muted)` } }, why[m.key] || ``)
        ]);
      })),
      elQ(NoteQ, { key: `n` }, `模型數字由年度模型拆分到季度；完整對照見「資金模型 → 各期收支 → 季度追蹤」。`)
    ]);
  }

  // 2｜Backlog 不是現金
  let mx2 = Math.max(b.scheduled, T.gross, T.newRev);
  S(`RPO 對照（不驅動營收）`, `RPO 只作對照：排程 RPO $${Y(b.scheduled, 0)}bn 五期可實現 $${Y(b.collected, 0)}bn，但同期要先投入 CapEx $${Y(T.gross, 0)}bn`, [
    elQ(BarQ, { key: 1, label: `排程 RPO（五期應認列）`, sub: `評價日 RPO＋期後新增，依季報桶分攤（只作對照）`, val: b.scheduled, max: mx2, color: `var(--color-accent)` }),
    elQ(BarQ, { key: 2, label: `扣：產能瓶頸`, sub: `排程 > 容量的部分收不到、不遞延`, val: b.bottleneck, max: mx2, color: `var(--color-bad)`, neg: !0 }),
    elQ(BarQ, { key: 3, label: `扣：信用損失`, val: b.credit, max: mx2, color: `var(--color-bad)`, neg: !0 }),
    elQ(BarQ, { key: 4, label: `可實現 RPO 收入`, val: b.collected, max: mx2, color: `var(--color-accent)` }),
    elQ(BarQ, { key: 5, label: `新簽約收入（尚未簽署）`, sub: `容量 > 排程的部分 × 新產能簽約率`, val: T.newRev, max: mx2, color: `var(--color-watch)` }),
    elQ(BarQ, { key: 6, label: `模型期毛 CapEx`, sub: `MW 連動，每 MW $${Y(cost30, 0)}m`, val: T.gross, max: mx2, color: `var(--color-ink)` }),
    elQ(`p`, { key: `n`, style: { fontSize: 15.5, color: `var(--color-muted)`, marginTop: `auto`, lineHeight: 1.55 } },
      `以上為營收，不是現金利潤：營收乘上 EBITDAR 率才是營運現金，還要再付租金 $${Y(T.lease, 0)}bn、利息 $${Y(T.interest, 0)}bn（含新融資利息）。`)
  ]);

  // 3｜單位經濟
  let mx3 = Math.max(ebMW, recov) * 1.1;
  S(`單位經濟`, `FY30 每 MW 每年 EBITDA $${Y(ebMW, 1)}m，${ebMW < recov ? `低於` : `高於`} GPU 年化資本回收 $${Y(recov, 1)}m`, [
    elQ(BarQ, { key: 1, label: `每 MW 年 EBITDA`, sub: `$${Y(rev30, 1)}m 年收入 × ${hA(util30 * 100, 0)} 利用率 × ${hA(eb30 * 100, 0)} EBITDA 率`, val: ebMW, max: mx3, color: `var(--color-accent)`, fmt: x => `$${Y(x, 1)}m` }),
    elQ(BarQ, { key: 2, label: `GPU 年化資本回收`, sub: `$${Y(cost30, 0)}m 建置成本 × 回收係數 ${Y(crf, 3)}（WACC ${hA(o.wacc * 100, 0)}、${e.gpuLife} 年）`, val: recov, max: mx3, color: `var(--color-bad)`, fmt: x => `$${Y(x, 1)}m` }),
    elQ(`div`, { key: `g`, style: { display: `grid`, gridTemplateColumns: `repeat(3, 1fr)`, gap: 18, marginTop: 30 } }, [
      elQ(StatQ, { key: 1, label: `每 MW 每年差額`, value: `${ebMW - recov < 0 ? `−` : `+`}$${Y(Math.abs(ebMW - recov), 1)}m`, tone: ebMW < recov ? `var(--color-bad)` : `var(--color-ok)`, note: `擴張本身${ebMW < recov ? `不創造` : `創造`}價值` }),
      elQ(StatQ, { key: 2, label: `打平所需每 MW 年收入`, value: `$${Y(beRev, 1)}m`, note: `目前 $${Y(rev30, 1)}m（${beRev >= rev30 ? `+` : `−`}${hA(Math.abs(beRev / rev30 - 1) * 100, 0)}）` }),
      elQ(StatQ, { key: 3, label: `或打平所需每 MW 建置成本`, value: `$${Y(beCost, 1)}m`, note: `目前 $${Y(cost30, 0)}m（${beCost >= cost30 ? `+` : `−`}${hA(Math.abs(beCost / cost30 - 1) * 100, 0)}）` })
    ]),
    elQ(`p`, { key: `n`, style: { fontSize: 15.5, color: `var(--color-muted)`, marginTop: `auto`, lineHeight: 1.55 } },
      `EBITDA 率由季報 AI cloud 分部 ${hA(e.ebStart * 100, 0)} 線性變動至 FY30 ${hA(e.ebSteady * 100, 0)}；取可觀察 neocloud 區間（IREN 約 35%、CRWV 約 59%），不取自每 MW 推導的加成。`)
  ]);

  // 4｜融資
  let mx4 = Math.max(...y.map(t => t.newDebt + t.convNew + t.equity + t.junk), .1);
  let nfi = y.reduce((a, t) => a + t.newDebtInt, 0);
  S(`融資`, `融資前缺口 $${Y(-Math.min(0, T.preFinEnd), 1)}bn：新債 $${Y(T.newDebt, 1)}bn、可轉債 $${Y(T.convNew, 1)}bn、股權 $${Y(T.equity, 1)}bn、高息債 $${Y(T.junk, 1)}bn`, [
    elQ(`div`, { key: `c`, style: { display: `grid`, gridTemplateColumns: `repeat(5, 1fr)`, gap: 28, alignItems: `end`, height: 320, padding: `0 20px` } }, y.map(t => {
      let tot = t.newDebt + t.convNew + t.equity + t.junk, H = 260;
      return elQ(`div`, { key: t.year, style: { display: `flex`, flexDirection: `column`, alignItems: `center`, justifyContent: `flex-end`, height: `100%` } }, [
        elQ(`div`, { key: `v`, style: { fontSize: 18, fontWeight: 700, marginBottom: 6, fontVariantNumeric: `tabular-nums` } }, `$${Y(tot, 1)}`),
        elQ(`div`, { key: `s`, style: { width: 96, display: `flex`, flexDirection: `column-reverse` } }, [
          [t.newDebt, `var(--color-accent)`], [t.convNew, `var(--color-ok)`], [t.equity, `var(--color-watch)`], [t.junk, `var(--color-bad)`]
        ].map(([x, c], i) => elQ(`div`, { key: i, style: { height: Math.max(0, x) / mx4 * H, background: c } }))),
        elQ(`div`, { key: `y`, style: { fontSize: 17, fontWeight: 600, marginTop: 10 } }, t.year),
        elQ(`div`, { key: `d`, style: { fontSize: 14, color: `var(--color-muted)`, marginTop: 2 } }, `總債務 $${Y(t.totalDebtEnd, 0)}bn`)
      ]);
    })),
    elQ(`div`, { key: `l`, style: { display: `flex`, gap: 26, fontSize: 15, marginTop: 36, paddingLeft: 20 } }, [
      [`新債（總債務 ≤ ${Y(e.debtBacklog, 1)}× backlog）`, `var(--color-accent)`],
      [`可轉債（每年 ≤ ${Y(e.cvCap ?? 0, 1)}bn）`, `var(--color-ok)`],
      [`股權（$${Y(e.eqPx, 2)} 折價 ${hA(e.eqDisc * 100, 0)}，每年上限${e.eqCapPct >= 9 ? `：無` : `＝現市值 ${hA(e.eqCapPct * 100, 0)}`}）`, `var(--color-watch)`],
      [`高息債 ${hA(e.junkRate * 100, 0)}`, `var(--color-bad)`]
    ].map(([a, c]) => elQ(`span`, { key: a, style: { display: `inline-flex`, alignItems: `center`, gap: 8 } }, [
      elQ(`i`, { key: `i`, style: { width: 14, height: 14, background: c, display: `inline-block`, borderRadius: 3 } }), a
    ]))),
    elQ(`p`, { key: `n`, style: { fontSize: 15.5, color: `var(--color-muted)`, marginTop: `auto`, lineHeight: 1.55 } },
      `新融資本身的利息 $${Y(nfi, 1)}bn 也由瀑布支應；每期在需要前先融足、期末現金不低於 $${Y(e.minCash ?? 2, 1)}bn。${eqYears.length ? `股權需求落在 ${eqYears.join('、')}，新股 ${Y(T.newShares, 2)}bn 股。` : `本情境不需股權。`}`)
  ]);

  // 5｜三情境
  let rowsQ = [
    [`${PERIODS[4]} 年底已連網 MW`, s => Y(SCENARIOS[s.sc].acc[4], 0)],
    [`每 MW 年收入（US$m/MW-IT）`, s => Y(SCENARIOS[s.sc].rev[4] * 1e3, 2)],
    [`模型期毛 CapEx（$bn）`, s => Y(s.capex, 0)],
    [`融資前缺口（$bn）`, s => Y(s.gap, 1)],
    [`新債／可轉債／股權／高息債（$bn）`, s => `${Y(s.nd, 1)}／${Y(s.cv, 1)}／${Y(s.eq, 1)}／${Y(s.jk, 1)}`],
    [`FY30 總債務（$bn）`, s => Y(s.debt30, 1)],
    [`FY30 營收／EBITDA（$bn）`, s => `${Y(s.rev30, 1)}／${Y(s.eb30, 1)}`],
    [`DCF 腿／EV/EBITDA 腿（每股）`, s => `$${Y(s.dcf, 1)}／$${Y(s.ev, 1)}`],
    [`加權目標價（0 截斷）`, s => `$${Y(s.tgt, 1)}`],
    [`加權目標價（選擇權模式）`, s => `$${Y(s.opt, 1)}`],
    [`結論`, s => s.call]
  ];
  S(`三情境`, `三個情境只改擴張力道，結論皆為${scen.every(s => s.call === scen[0].call) ? `「${scen[0].call}」` : `不同：${scen.map(s => s.call).join('／')}`}${scen[0].tgt > scen[1].tgt && scen[1].tgt > scen[2].tgt ? `；蓋得愈多，目標價愈低` : ``}`, [
    elQ(`table`, { key: `t`, style: { width: `100%`, borderCollapse: `collapse`, fontSize: 16.5, lineHeight: 1.35, fontVariantNumeric: `tabular-nums` } }, [
      elQ(`thead`, { key: `h` }, elQ(`tr`, {}, [elQ(`th`, { key: `x`, style: { textAlign: `left`, padding: `6px 12px`, borderBottom: `2px solid var(--color-ink)` } }, ``),
        ...scen.map(s => elQ(`th`, { key: s.sc, style: { textAlign: `right`, padding: `6px 12px`, borderBottom: `2px solid var(--color-ink)`, color: s.sc === e.scenario ? `var(--color-accent)` : `var(--color-fg)` } }, s.label))])),
      elQ(`tbody`, { key: `b` }, rowsQ.map(([lab, fn], i) => elQ(`tr`, { key: lab, style: { background: i >= 7 ? `var(--color-surface)` : `transparent`, fontWeight: i === 7 ? 700 : 400 } }, [
        elQ(`td`, { key: `l`, style: { padding: `6px 12px`, borderBottom: `1px solid var(--color-border)` } }, lab),
        ...scen.map(s => elQ(`td`, { key: s.sc, style: { padding: `6px 12px`, textAlign: `right`, borderBottom: `1px solid var(--color-border)`, color: i === 9 ? (s.call === `賣出` ? `var(--color-bad)` : s.call === `買進` ? `var(--color-ok)` : `var(--color-watch)`) : `inherit` } }, fn(s)))
      ])))
    ]),
    elQ(`p`, { key: `j`, style: { fontSize: 16, lineHeight: 1.5, margin: `16px 0 0` } },
      `情境區間（保守與積極情境）$${Y(TR.A[0], 1)}–$${Y(TR.A[1], 1)}。${TR.judge}`),
    elQ(`p`, { key: `n`, style: { fontSize: 14, color: `var(--color-muted)`, marginTop: `auto`, lineHeight: 1.5 } },
      `加權目標價＝DCF ${hA((f.call.weights?.dcf ?? .45) * 100, 0)}＋EV/EBITDA（${PERIOD_LABELS[o.evYear ?? 1]}，${Y(o.evEbitda, 1)}x）${hA((f.call.weights?.pe ?? .55) * 100, 0)}。DCF 股權價值為負時以 0 截斷；選擇權模式以 Merton（σ ${hA(o.sigma * 100, 0)}）估計有限責任下的股權價值。`)
  ]);

  // 6｜反向 DCF
  S(`反向 DCF`, rv && Number.isFinite(rv.R)
    ? `現價 $${Y(P, 2)} 要成立：FY30 每 MW 年收入需 ${rv.R >= 1 ? `+` : `−`}${hA(Math.abs(rv.R - 1) * 100, 0)}${Number.isFinite(rv.C) ? `，或建置成本需 ${rv.C >= 1 ? `+` : `−`}${hA(Math.abs(rv.C - 1) * 100, 0)}` : `；建置成本單獨調整無解`}`
    : `現價要成立需要什麼（計算中或無解）`, rv ? [
      elQ(`div`, { key: `g`, style: { display: `grid`, gridTemplateColumns: `repeat(3, 1fr)`, gap: 20 } }, [
        elQ(StatQ, { key: 1, label: `每 MW 年收入（FY30）`, value: Number.isFinite(rv.R) ? `$${Y(rv.rev30 * rv.R, 1)}m` : `無解`, note: `目前 $${Y(rv.rev30, 1)}m；其他條件不變` }),
        elQ(StatQ, { key: 2, label: `或每 MW 建置成本`, value: Number.isFinite(rv.C) ? `$${Y(rv.cost30 * rv.C, 1)}m` : `無解`, note: `目前 $${Y(rv.cost30, 0)}m；其他條件不變` }),
        elQ(StatQ, { key: 3, label: `或穩態 EBITDA 率`, value: Number.isFinite(rv.Eb) ? hA(rv.Eb * 100, 0) : `無解`, tone: Number.isFinite(rv.Eb) && rv.Eb > .59 ? `var(--color-bad)` : void 0, note: `目前 ${hA(rv.eb30 * 100, 0)}；可觀察 neocloud 上緣約 59%（CRWV）` })
      ]),
      elQ(`div`, { key: `m`, style: { display: `grid`, gridTemplateColumns: `1fr 1fr`, gap: 36, marginTop: 22, alignItems: `start` } }, [
        elQ(`table`, { key: `t`, style: { borderCollapse: `collapse`, fontSize: 16, fontVariantNumeric: `tabular-nums`, width: `100%` } }, [
          elQ(`thead`, { key: `h` }, elQ(`tr`, {}, [elQ(`th`, { key: `x`, style: { textAlign: `left`, padding: `5px 10px`, borderBottom: `2px solid var(--color-ink)`, fontWeight: 600 } }, `建置成本 ＼ EBITDA 率`),
            ...rv.ebs.map(q => elQ(`th`, { key: q, style: { textAlign: `right`, padding: `5px 10px`, borderBottom: `2px solid var(--color-ink)` } }, hA(q * 100, 0)))])),
          elQ(`tbody`, { key: `b` }, rv.caps.map((c, i) => elQ(`tr`, { key: c }, [
            elQ(`td`, { key: `l`, style: { padding: `5px 10px`, borderBottom: `1px solid var(--color-border)` } }, `$${Y(rv.cost30 / (e.capexScale ?? 1) * c, 1)}m（${hA(c * 100, 0)}）`),
            ...rv.grid[i].map((x, j) => elQ(`td`, { key: j, style: { padding: `5px 10px`, textAlign: `right`, borderBottom: `1px solid var(--color-border)`, background: Number.isFinite(x) && x * rv.rev30 / (e.revScale ?? 1) <= rv.rev30 ? `var(--color-ok-bg)` : `transparent` } },
              Number.isFinite(x) ? `$${Y(x * rv.rev30 / (e.revScale ?? 1), 1)}m` : `—`))
          ])))
        ]),
        elQ(`p`, { key: `p`, style: { fontSize: 16.5, lineHeight: 1.65, margin: 0 } },
        `解法：固定其他輸入，只調一個變數，使 DCF 每股＝現價。${Number.isFinite(rv.Rt) ? `若改以加權目標價＝現價反解，每 MW 年收入需 ${rv.Rt >= 1 ? `+` : `−`}${hA(Math.abs(rv.Rt - 1) * 100, 0)}。` : ``}左表為 DCF＝現價所需的 FY30 每 MW 年收入（綠底＝不高於目前 $${Y(rv.rev30, 1)}m）。`)
      ]),
      elQ(`p`, { key: `n`, style: { fontSize: 15, color: `var(--color-muted)`, marginTop: `auto`, lineHeight: 1.55 } },
        Number.isFinite(rv.Eb) && rv.Eb > .59 ? `解讀：現價隱含單位經濟大幅改善（更高租價或更低 GPU 成本）；單靠營運效率，EBITDA 率須超出可觀察 neocloud 上緣。` : `解讀：現價所需條件落在可觀察 neocloud 範圍內，需逐項檢驗。`)
    ] : elQ(`p`, {}, `計算中…`));

  // 7｜評價方法：錨定年度 × 倍數（v3.5）
  {
    let k6 = eg.mults.indexOf(6), best = { v: -1 };
    eg.leg.forEach((r, i) => r.forEach((x, j) => { if (x > best.v) best = { v: x, m: eg.mults[i], y: eg.years[j] }; }));
    let hits = [];
    eg.leg.forEach((r, i) => r.forEach((x, j) => { if (x >= .95 * P) hits.push(`${eg.years[j]} × ${Y(eg.mults[i], 1)}x（$${Y(x, 1)}）`); }));
    let tmax = Math.max(...eg.tgt.flat()), l29 = k6 >= 0 ? eg.leg[k6][2] : NaN;
    S(`評價方法`, `結論方向不變，但幅度取決於錨定年度與倍數：${Number.isFinite(l29) ? `FY29 × 6x 時 EV/EBITDA 腿為 $${Y(l29, 1)}，${l29 >= .95 * P ? `約等於` : l29 > P ? `高於` : `低於`}現價` : `見下表`}`, [
      elQ(EvGridQ, { key: `g`, st: e, o: o, g: eg, big: !0 }),
      elQ(`div`, { key: `n`, style: { marginTop: 22, fontSize: 17, lineHeight: 1.6 } }, [
        elQ(`div`, { key: 1 }, `• 矩陣內加權目標價最高 $${Y(tmax, 1)}，${tmax < P ? `仍低於現價，賣出方向在所有組合下成立` : `部分組合高於現價，結論對方法選擇敏感`}。`),
        elQ(`div`, { key: 2 }, hits.length ? `• 但 EV/EBITDA 腿單獨達到現價 95% 以上的組合：${hits.join('、')}——市場大致以穩態倍數上緣定價 FY28 以後的 EBITDA。` : `• EV/EBITDA 腿在所有組合下都低於現價的 95%。`),
        elQ(`div`, { key: 3 }, `• 因此結論實際依賴：(a) DCF 腿（權重 ${hA(eg.wd * 100, 0)} 取 $${Y(eg.dcf, 1)}）；(b) 倍數 ${Y(o.evEbitda, 1)}x 沿用模板的穩態上緣（未另估 ${COMPANY_DATA.meta.company}）。`)
      ]),
      elQ(`p`, { key: `f`, style: { fontSize: 14, color: `var(--color-muted)`, marginTop: `auto`, lineHeight: 1.5 } },
        `目前設定：錨定 ${PERIOD_LABELS[o.evYear ?? 1]}、${Y(o.evEbitda, 1)}x（黃底）。FY27 錨定的疑慮：6x 為穩態倍數卻套在爬坡年度；FY27 年末淨負債已含 FY28 才產生 EBITDA 的預建 CapEx。`)
    ]);
  }

  // 7｜敏感度
  let top = m.slice(0, 7), mx7 = Math.max(...top.map(t => t.high), tgt) * 1.08;
  S(`敏感度`, `${top.every(t => t.high < P) ? `任一單一變數在測試範圍內，都推不到現價 $${Y(P, 2)}` : `部分單一變數在測試範圍內可推到現價 $${Y(P, 2)}`}；最敏感的是${top[0].name}與${top[1].name}`, [
    elQ(`div`, { key: `t`, style: { position: `relative` } }, [
      ...top.map(t => elQ(`div`, { key: t.name, style: { display: `grid`, gridTemplateColumns: `230px 1fr`, alignItems: `center`, gap: 16, padding: `7px 0` } }, [
        elQ(`div`, { key: `l`, style: { fontSize: 18, fontWeight: 600 } }, [t.name, elQ(`div`, { key: `s`, style: { fontSize: 13.5, fontWeight: 400, color: `var(--color-muted)` } }, `${t.la} ↔ ${t.lb}`)]),
        elQ(`div`, { key: `b`, style: { position: `relative`, height: 34 } }, [
          elQ(`div`, { key: `r`, style: { position: `absolute`, left: `${t.low / mx7 * 100}%`, width: `${Math.max(.4, (t.high - t.low) / mx7 * 100)}%`, top: 4, bottom: 4, background: `var(--color-accent)`, opacity: .85, borderRadius: 3 } }),
          elQ(`div`, { key: `a`, style: { position: `absolute`, left: `calc(${t.low / mx7 * 100}% - 58px)`, width: 52, textAlign: `right`, top: 6, fontSize: 15, fontVariantNumeric: `tabular-nums` } }, `$${Y(t.low, 1)}`),
          elQ(`div`, { key: `c`, style: { position: `absolute`, left: `calc(${t.high / mx7 * 100}% + 6px)`, top: 6, fontSize: 15, fontVariantNumeric: `tabular-nums` } }, `$${Y(t.high, 1)}`),
          elQ(`div`, { key: `m`, style: { position: `absolute`, left: `${tgt / mx7 * 100}%`, top: 0, bottom: 0, width: 2, background: `var(--color-ink)` } })
        ])
      ]))
    ]),
    elQ(`p`, { key: `n`, style: { fontSize: 15, color: `var(--color-muted)`, marginTop: `auto`, lineHeight: 1.55 } },
      `黑線為目前目標價 $${Y(tgt, 1)}；各列只改一個變數。完整清單（${m.length} 項）見「資金模型 → 分析與檢查 → 敏感性」。`)
  ]);

  // 8｜驗證點與限制
  let li = (a, k) => elQ(`li`, { key: k, style: { margin: `0 0 10px`, lineHeight: 1.55 } }, a);
  S(`驗證點與限制`, `下一個驗證點是 ${CALL_FACTS.nextEarn || `下一季財報`}`, [
    elQ(`div`, { key: `g`, style: { display: `grid`, gridTemplateColumns: `1fr 1fr`, gap: 40 } }, [
      elQ(`div`, { key: `a` }, [
        elQ(`div`, { key: `h`, style: { fontSize: 20, fontWeight: 700, marginBottom: 10 } }, `會改變結論的觀察值`),
        elQ(`ul`, { key: `u`, style: { fontSize: 18, paddingLeft: 24, margin: 0, listStyle: `disc` } }, [
          li(`新簽約的每 MW 單價：若持續高於 $${rv && Number.isFinite(rv.R) ? Y(rv.rev30 * rv.R, 1) : `—`}m（反向 DCF 門檻），單位經濟翻正`, 1),
          li(`調整後營業利益率與 EBITDA 率路徑：模型假設 ${hA(e.ebStart * 100, 0)} → ${hA(e.ebSteady * 100, 0)}`, 2),
          li(`新債利率：模型 ${rateLo === rateHi ? Y(rateLo, 1) : `${Y(rateLo, 1)}–${Y(rateHi, 1)}`}%；利差擴大會壓縮債務容量`, 3),
          li(`D&A／營收：GPU 經濟壽命（模型 ${e.gpuLife} 年）是否被延長或縮短`, 4)
        ])
      ]),
      elQ(`div`, { key: `b` }, [
        elQ(`div`, { key: `h`, style: { fontSize: 20, fontWeight: 700, marginBottom: 10 } }, `模型限制`),
        elQ(`ul`, { key: `u`, style: { fontSize: 18, paddingLeft: 24, margin: 0, listStyle: `disc` } }, [
          li(`DCF 股權價值${f.d.perShareRaw < 0 ? `為負（未截斷每股 −$${Y(-f.d.perShareRaw, 1)}），以 0 截斷` : `為正`}；截斷時，只影響 DCF 的變數（如股權折價與上限）不反映在目標價`, 1),
          li(`終值現值 ${mA(f.d.pvTv, 1)}bn、五期 FCF 現值 ${mA(f.d.pvFcf, 1)}bn：企業價值幾乎全來自終值，DCF 對 WACC、永續成長極敏感`, 2),
          li(`租金尚未改為 MW 驅動；以具名站點基準計，五期可能${rentGap >= 0 ? `低估` : `高估`}約 $${Y(Math.abs(rentGap), 1)}bn`, 3),
          li(`EV/EBITDA 腿錨定 ${PERIOD_LABELS[o.evYear ?? 1]}；錨定年度與倍數的選擇會大幅改變目標價（見評價方法頁）`, 4)
        ])
      ])
    ]),
    elQ(`p`, { key: `n`, style: { fontSize: 15, color: `var(--color-muted)`, marginTop: `auto` } }, `本頁為研究框架摘要，不是投資建議。`)
  ]);

  // 最後一頁｜一頁摘要（v4.3 起；v0.2 由第 1 頁移到最後）：結論、與市場的差異、現價隱含什麼、驗證點
  S(`一頁摘要`, `${f.call.call}：點位 $${Y(TR.pt, 1)}（情境區間 $${Y(TR.A[0], 1)}–$${Y(TR.A[1], 1)}）；${cv.first < 0 ? `與市場共識差距在 ${pctQ(CONS_TOL)} 以內` : `與市場共識的分歧始於 ${CONS_YEARS[cv.first]}`}`, onePageQ({ cv, qv, TR, f, o, e, rv, scLabel, callTone }), 28);

  let N = slides.length, mk = (s, i) => elQ(SlideQ, { no: i + 1, total: N, kicker: s.kicker, title: s.title, foot, tsize: s.tsize, main: s.main }, s.body);

  // 簡報模式：鍵盤 ←／→／空白鍵換頁，Esc 離開
  (0, v.useEffect)(() => {
    if (!pres) return;
    let h = ev => {
      if (ev.key === `ArrowRight` || ev.key === `PageDown` || ev.key === ` `) { ev.preventDefault(); si(i => Math.min(N - 1, i + 1)); }
      else if (ev.key === `ArrowLeft` || ev.key === `PageUp`) { ev.preventDefault(); si(i => Math.max(0, i - 1)); }
      else if (ev.key === `Escape`) sp(!1);
    };
    window.addEventListener(`keydown`, h);
    return () => window.removeEventListener(`keydown`, h);
  }, [pres, N]);
  let [vw, svw] = (0, v.useState)(() => [window.innerWidth, window.innerHeight]);
  (0, v.useEffect)(() => {
    let h = () => svw([window.innerWidth, window.innerHeight]);
    window.addEventListener(`resize`, h);
    return () => window.removeEventListener(`resize`, h);
  }, []);
  let open = () => {
    sp(!0); si(0);
    try { document.documentElement.requestFullscreen?.() } catch {}
  }, close = () => {
    sp(!1);
    try { document.fullscreenElement && document.exitFullscreen?.() } catch {}
  };
  let ks = Math.min(vw[0] / SW_Q, (vw[1] - 56) / SH_Q);

  return elQ(`div`, { className: `sumQ-wrap`, style: { display: `flex`, flexDirection: `column`, gap: 18 } }, [
    elQ(`style`, { key: `css` }, SUMCSSQ),
    elQ(`div`, { key: `bar`, className: `sumQ-toolbar`, style: { display: `flex`, flexWrap: `wrap`, alignItems: `center`, gap: 10 } }, [
      elQ(`button`, { key: `p`, className: `sumQ-btn pri`, onClick: open }, `簡報模式（全螢幕）`),
      elQ(`button`, { key: `r`, className: `sumQ-btn`, onClick: () => window.print() }, `列印／另存 PDF`),
      elQ(`span`, { key: `t`, style: { fontSize: 13.5, color: `var(--color-muted)` } },
        `${N} 頁，數字即時連動目前輸入（${scLabel}）。簡報模式以 ←／→ 換頁、Esc 離開；列印時請選「橫向」、邊界「無」、勾選「背景圖形」。`)
    ]),
    ...slides.map((s, i) => elQ(ScaleQ, { key: i }, mk(s, i))),
    pres ? elQ(`div`, {
      key: `pres`, ref: pr, className: `sumQ-pres`,
      style: { position: `fixed`, inset: 0, zIndex: 9999, background: `var(--color-ink)`, display: `flex`, flexDirection: `column`, alignItems: `center`, justifyContent: `center` }
    }, [
      elQ(`div`, { key: `s`, style: { width: SW_Q * ks, height: SH_Q * ks } },
        elQ(`div`, { key: ix, style: { width: SW_Q, height: SH_Q, transform: `scale(${ks})`, transformOrigin: `top left` } }, mk(slides[ix], ix))),
      elQ(`div`, { key: `c`, style: { display: `flex`, gap: 10, alignItems: `center`, height: 48, color: `var(--color-accent-fg)`, fontSize: 14 } }, [
        elQ(`button`, { key: `a`, className: `sumQ-btn`, onClick: () => si(i => Math.max(0, i - 1)), disabled: ix === 0 }, `上一頁`),
        elQ(`span`, { key: `n`, style: { minWidth: 60, textAlign: `center` } }, `${ix + 1} / ${N}`),
        elQ(`button`, { key: `b`, className: `sumQ-btn`, onClick: () => si(i => Math.min(N - 1, i + 1)), disabled: ix === N - 1 }, `下一頁`),
        elQ(`button`, { key: `x`, className: `sumQ-btn`, onClick: close }, `離開（Esc）`)
      ])
    ]) : null
  ]);
}
