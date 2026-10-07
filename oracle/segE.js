// ===== segE.js：「總結」簡報頁（v3.4 新增）=====
// 所有數字即時取自資金引擎（dA）與評價引擎（TM），不另存常數；切換情境或修改輸入後自動更新。
// 版面固定為 1280×720 的投影片，一般檢視時等比縮放；簡報模式全螢幕逐頁播放；列印時每頁一張（16:9）。
// 全域名稱一律以 Q 結尾，避開模板函式庫。

var SW_Q = 1280, SH_Q = 720;

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

// 單張投影片外框：主訊息標題＋內容＋頁尾
function SlideQ({ no, total, kicker, title, children, foot, tsize }) {
  return elQ(`section`, {
    className: `sumQ-slide`,
    style: {
      width: SW_Q, height: SH_Q, boxSizing: `border-box`, padding: `44px 56px 30px`,
      background: `var(--color-card)`, color: `var(--color-fg)`, border: `1px solid var(--color-border)`,
      borderRadius: 14, display: `flex`, flexDirection: `column`, overflow: `hidden`, fontFamily: `var(--font-sans)`
    }
  }, [
    elQ(`div`, { key: `k`, style: { fontSize: 15, color: `var(--color-accent)`, fontWeight: 600 } }, kicker),
    elQ(`h2`, { key: `t`, style: { margin: `8px 0 0`, fontSize: tsize || 34, lineHeight: 1.3, fontWeight: 700, letterSpacing: `-.01em`, maxWidth: 1160, textWrap: `pretty` } }, title),
    elQ(`div`, { key: `b`, style: { flex: 1, minHeight: 0, marginTop: tsize ? 16 : 26, display: `flex`, flexDirection: `column` } }, children),
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
      ? `反向 DCF（DCF＝現價，其他不變）：FY30 每 MW 年收入需 $${Y(rv.rev30 * rv.R, 1)}m（${rv.R >= 1 ? `+` : `−`}${hA(Math.abs(rv.R - 1) * 100, 0)}），或建置成本 $${Y(rv.cost30 * rv.C, 1)}m（${rv.C >= 1 ? `+` : `−`}${hA(Math.abs(rv.C - 1) * 100, 0)}），或穩態 EBITDA 率 ${Number.isFinite(rv.Eb) ? hA(rv.Eb * 100, 0) : `無解`}${Number.isFinite(rv.Rt) ? `；加權目標價＝現價需每 MW 年收入 ${rv.Rt >= 1 ? `+` : `−`}${hA(Math.abs(rv.Rt - 1) * 100, 0)}` : ``}。`
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
.sumQ-btn { height: 40px; padding: 0 16px; border-radius: 8px; font-size: 14px; font-weight: 600; cursor: pointer; border: 1px solid var(--color-border); background: var(--color-card); color: var(--color-fg); }
.sumQ-btn:hover { background: var(--color-surface); }
.sumQ-btn:focus-visible { outline: 2px solid var(--color-accent); outline-offset: 2px; }
.sumQ-btn.pri { background: var(--color-ink); color: var(--color-accent-fg); border-color: var(--color-ink); }
`;

function SumQ({ d, f, e, o, m, tr: TR, active }) {
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
  // --- 反向 DCF（約 0.4 秒，只在本頁顯示時計算）---
  let rv = (0, v.useMemo)(() => active ? reverseDcf(e, o) : null, [e, o, active]);
  let cv = (0, v.useMemo)(() => consensusView(d, f, o, TR, e), [d, f, o, TR, e]), qv = (0, v.useMemo)(() => quarterlyView(d, f, e), [d, f, e]);
  let eg = (0, v.useMemo)(() => evAnchorGrid(d, e, o), [d, e, o]);
  let [pres, sp] = (0, v.useState)(!1), [ix, si] = (0, v.useState)(0), pr = (0, v.useRef)(null);

  let y = d.years, T = d.totals, b = rpoBridge(d), P = o.price, tgt = f.call.blended, up = f.call.upside;
  let scLabel = e.scenario === `custom` ? `自訂情境` : SCENARIOS[e.scenario]?.label;
  let ver = VLOG[VLOG.length - 1][0];
  let foot = `${COMPANY_DATA.meta.company} 收支模型 ${ver}（${UPDATE_DATE}）· ${scLabel} · 現價 $${Y(P, 2)}（${COMPANY_DATA.meta.priceDate} 收盤）· 研究框架，不是投資建議`;
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
  let S = (kicker, title, body, tsize) => slides.push({ kicker: slides.length ? `附錄｜${kicker}` : kicker, title, body, tsize });

  // 0｜一頁摘要（v4.3）：結論、與市場的差異、現價隱含什麼、驗證點；原 9 頁順延為附錄
  S(`一頁摘要`, `${f.call.call}：點位 $${Y(TR.pt, 1)}（情境區間 $${Y(TR.A[0], 1)}–$${Y(TR.A[1], 1)}）；${cv.first < 0 ? `與市場共識差距在 ${pctQ(CONS_TOL)} 以內` : `與市場共識的分歧始於 ${CONS_YEARS[cv.first]}`}`, onePageQ({ cv, qv, TR, f, o, e, rv, scLabel, callTone }), 28);

  // 1｜結論
  S(`結論`, `${TR.head}，較現價 $${Y(P, 2)} ${up >= 0 ? `高` : `低`} ${hA(Math.abs(up) * 100, 0)}：${f.call.call}`, [
    elQ(`p`, { key: `p`, style: { fontSize: 21, lineHeight: 1.6, color: `var(--color-muted)`, margin: 0, maxWidth: 1080 } },
      `${COMPANY_DATA.texts.thesis}：營收由已連網 MW × 每 MW 年收入驅動，客戶預付在建置前先收現；問題是預付能覆蓋多少資本支出、剩下的缺口要靠多少可轉債與新股。`),
    elQ(`div`, { key: `g`, style: { display: `grid`, gridTemplateColumns: `repeat(3, 1fr)`, gap: 18, marginTop: 26 } }, scen.map(s => elQ(`div`, {
      key: s.sc, style: { padding: `18px 22px`, borderRadius: 10, background: s.sc === e.scenario ? `var(--color-ink)` : `var(--color-surface)`, color: s.sc === e.scenario ? `var(--color-accent-fg)` : `var(--color-fg)` }
    }, [
      elQ(`div`, { key: `a`, style: { fontSize: 17, fontWeight: 600, opacity: .85 } }, s.label),
      elQ(`div`, { key: `b`, style: { fontSize: 46, fontWeight: 700, marginTop: 4, fontVariantNumeric: `tabular-nums` } }, `$${Y(s.tgt, 1)}`),
      elQ(`div`, { key: `c`, style: { fontSize: 15, marginTop: 4, opacity: .8 } }, `選擇權模式 $${Y(s.opt, 1)} · ${s.call}`)
    ]))),
    elQ(`p`, { key: `j`, style: { fontSize: 16, lineHeight: 1.5, margin: `14px 0 0`, maxWidth: 1180 } },
      `${TR.judge}${TR.bLabel} $${Y(TR.B[0], 1)}–$${Y(TR.B[1], 1)}：${TR.pos}；評等依點位。`),
    elQ(`div`, { key: `r`, style: { display: `grid`, gridTemplateColumns: `repeat(3, 1fr)`, gap: 18, marginTop: 22 } }, [
      [`收入受產能約束`, `排程 RPO $${Y(b.scheduled, 1)}bn，五期可實現 $${Y(b.collected, 1)}bn；模型期毛 CapEx $${Y(T.gross, 0)}bn。`],
      [ebMW < recov ? `單位經濟為負` : `單位經濟為正`, `FY30 每 MW 年 EBITDA $${Y(ebMW, 1)}m，${ebMW < recov ? `低於` : `高於`} GPU 年化資本回收 $${Y(recov, 1)}m。`],
      [`依賴外部資金`, `融資前缺口 $${Y(T.preFinEnd < 0 ? -T.preFinEnd : 0, 1)}bn，FY30 總債務 $${Y(y[4].totalDebtEnd, 0)}bn。`]
    ].map(([a, c]) => elQ(`div`, { key: a, style: { borderTop: `3px solid var(--color-accent)`, paddingTop: 12 } }, [
      elQ(`div`, { key: `a`, style: { fontSize: 18, fontWeight: 700 } }, a),
      elQ(`div`, { key: `c`, style: { fontSize: 15.5, lineHeight: 1.55, color: `var(--color-muted)`, marginTop: 6 } }, c)
    ])))
  ]);

  // 2｜Backlog 不是現金
  let mx2 = Math.max(b.scheduled, T.gross, T.newRev);
  S(`RPO 對照與資本支出`, `排程 RPO $${Y(b.scheduled, 0)}bn 五期可實現 $${Y(b.collected, 0)}bn，但同期要先投入 CapEx $${Y(T.gross, 0)}bn`, [
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
      `EBITDA 率由 ${TXQ.ebStartSource} ${hA(e.ebStart * 100, 0)} 線性變動至 ${PERIODS[4]} ${hA(e.ebSteady * 100, 0)}；取可觀察 neocloud 區間（IREN 約 35%、CRWV 約 59%），不取自每 MW 推導的加成。`)
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
    elQ(`table`, { key: `t`, style: { width: `100%`, borderCollapse: `collapse`, fontSize: 17, fontVariantNumeric: `tabular-nums` } }, [
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
    ? `現價 $${Y(P, 2)} 要成立：FY30 每 MW 年收入需 ${rv.R >= 1 ? `+` : `−`}${hA(Math.abs(rv.R - 1) * 100, 0)}，或建置成本需 ${rv.C >= 1 ? `+` : `−`}${hA(Math.abs(rv.C - 1) * 100, 0)}`
    : `現價要成立需要什麼（計算中或無解）`, rv ? [
      elQ(`div`, { key: `g`, style: { display: `grid`, gridTemplateColumns: `repeat(3, 1fr)`, gap: 20 } }, [
        elQ(StatQ, { key: 1, label: `每 MW 年收入（${PERIODS[4]}）`, value: Number.isFinite(rv.R) ? `$${Y(rv.rev30 * rv.R, 1)}m` : `無解`, note: `目前 $${Y(rv.rev30, 1)}m；其他條件不變` }),
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
        elQ(`div`, { key: 3 }, `• 因此賣出結論實際依賴：(a) DCF 股權價值為負（權重 ${hA(eg.wd * 100, 0)} 取 $${Y(eg.dcf, 1)}）；(b) 6x 是穩態倍數上緣（約 3.4–6.0x）而非中值。`)
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
          li(`租金尚未改為 MW 驅動；以具名站點基準計，五期可能低估約 $${Y(rentGap, 1)}bn`, 3),
          li(`EV/EBITDA 腿錨定 ${PERIOD_LABELS[o.evYear ?? 1]}；錨定年度與倍數的選擇會大幅改變目標價（見評價方法頁）`, 4)
        ])
      ])
    ]),
    elQ(`p`, { key: `n`, style: { fontSize: 15, color: `var(--color-muted)`, marginTop: `auto` } }, `本頁為研究框架摘要，不是投資建議。`)
  ]);

  let N = slides.length, mk = (s, i) => elQ(SlideQ, { no: i + 1, total: N, kicker: s.kicker, title: s.title, foot, tsize: s.tsize }, s.body);

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
        elQ(`div`, { style: { width: SW_Q, height: SH_Q, transform: `scale(${ks})`, transformOrigin: `top left` } }, mk(slides[ix], ix))),
      elQ(`div`, { key: `c`, style: { display: `flex`, gap: 10, alignItems: `center`, height: 48, color: `var(--color-accent-fg)`, fontSize: 14 } }, [
        elQ(`button`, { key: `a`, className: `sumQ-btn`, onClick: () => si(i => Math.max(0, i - 1)), disabled: ix === 0 }, `上一頁`),
        elQ(`span`, { key: `n`, style: { minWidth: 60, textAlign: `center` } }, `${ix + 1} / ${N}`),
        elQ(`button`, { key: `b`, className: `sumQ-btn`, onClick: () => si(i => Math.min(N - 1, i + 1)), disabled: ix === N - 1 }, `下一頁`),
        elQ(`button`, { key: `x`, className: `sumQ-btn`, onClick: close }, `離開（Esc）`)
      ])
    ]) : null
  ]);
}
