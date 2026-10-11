function AM({
  result: e,
  state: t,
  v: n,
  pack: r,
  tr: TR,
  onPatch: i,
  tab: a,
  onTab: o
}) {
  let {
    fwd: s,
    d: c,
    peRaw: d,
    peAdj: f,
    hole: p,
    call: m,
    shares: h
  } = r, g = (0, v.useMemo)(() => dcfGridWaccG(s, r.v), [s, r.v]), RV = (0, v.useMemo)(() => a === 4 ? reverseDcf(t, n) : null, [t, n, a]), y = n.price * h, b = y + r.v.netDebt, NDT = n.netDebt + CVN.reduce((a, c) => a + (c.mand ? 0 : c.M), 0) + ndAdjQ(n), // v0.1b：模型列用評價股數與評價淨負債（含可轉債分類）；TTM 列市值為基本股，可轉債全數計入淨負債
 x = s[1].revenue, S = s[1].ebitda, C = S > 0 ? b / S : 0, E = [`損益簡表`, `Comps`, `DCF`, `目標價`], D = m.call === `買進` ? `ok` : m.call === `賣出` ? `bad` : `watch`, O = e.totals.end, Rt = (e, t) => (0, $.jsx)(`td`, {
    className: `py-1.5 text-right font-mono tabular-nums`,
    children: Number.isFinite(e) ? Y(e, t) : `—`
  });
  return (0, $.jsxs)(`div`, {
    className: `space-y-4`,
    children: [(0, $.jsxs)(Jj, {
      className: Vj(`border px-4 py-3`, m.funded ? `border-ok/30 bg-ok/5` : `border-bad/30 bg-bad/5`),
      children: [(0, $.jsxs)(`div`, {
        className: `flex flex-wrap items-start justify-between gap-3`,
        children: [(0, $.jsxs)(`div`, {
          className: `min-w-0`,
          children: [(0, $.jsx)(`div`, {
            className: `text-xs font-medium text-muted`,
            children: `資金模型 → 評價 連動`
          }), (0, $.jsxs)(`p`, {
            className: `mt-1 text-sm leading-relaxed`,
            children: [`算力產能、Cash CapEx、利息（含瀑布新債）、股數（含瀑布新股）皆來自資金模型。缺口在需要前一期先融好：依序動用額度、資產層新債（受 債務／backlog 上限）、股權；估值不再扣缺口本金。五期需股權 `, Y(p, 0), `bn、新股 `, Y(e.totals.newShares, 2), `bn 股。`]
          })]
        }), (0, $.jsxs)(`div`, {
          className: `flex shrink-0 flex-row flex-wrap items-center gap-2`,
          children: [(0, $.jsx)(Uj, {
            tone: m.funded ? `ok` : `bad`,
            children: m.funded ? `無需新股` : m.blocked ? `股權需求過大` : `需股權募資`
          }), (0, $.jsx)(Uj, {
            tone: D,
            children: m.call
          })]
        })]
      }), (0, $.jsx)(`div`, {
        className: `mt-3 grid grid-cols-2 gap-2 md:grid-cols-5`,
        children: [
          [`${PERIODS[1]} 算力產能`, mA(s[1].gpu) + `bn`],
          [`五期 Cash CapEx`, mA(e.totals.cashCapex) + `bn`],
          [`需股權募資`, mA(p) + `bn`],
          [`ATM／股數`, `${t.includeAtm?`含`:`不含`} · ${Y(h,3)}bn`],
          [`評價結論`, m.call]
        ].map(([e, t]) => (0, $.jsxs)(`div`, {
          className: `rounded-md bg-card/80 px-3 py-2`,
          children: [(0, $.jsx)(`div`, {
            className: `text-[11px] text-muted`,
            children: e
          }), (0, $.jsx)(`div`, {
            className: Vj(`mt-0.5 font-mono text-sm tabular-nums`, e === `需股權募資` && (p > .01 ? `text-bad` : `text-ok`), e === `評價結論` && (D === `ok` ? `text-ok` : D === `bad` ? `text-bad` : `text-watch`)),
            children: t
          })]
        }, e))
      })]
    }), (0, $.jsx)(`div`, {
      className: `grid grid-cols-2 gap-3 md:grid-cols-3 lg:grid-cols-6`,
      children: [
        [`現價（${mdQ(PRICE_DATE)} 收）`, `$${Y(n.price,2)}`],
        [`市值`, `${Y(y,0)}bn`],
        [`企業價值`, `${Y(b,0)}bn`],
        [`本模型目標價（情境區間）`, `$${Y(m.blended,1)}（$${Y(TR.A[0],1)}–$${Y(TR.A[1],1)}）`],
        [`潛在空間`, hA(m.upside * 100, 0)],
        [`投資結論`, m.call]
      ].map(([e, t], n) => (0, $.jsxs)(Jj, {
        className: `p-3 md:p-4`,
        children: [(0, $.jsx)(`div`, {
          className: `text-xs text-muted`,
          children: e
        }), (0, $.jsx)(`div`, {
          className: Vj(`mt-1 font-display text-xl font-semibold tabular-nums`, n === 4 && (m.upside >= 0 ? `text-ok` : `text-bad`), n === 5 && (D === `ok` ? `text-ok` : D === `bad` ? `text-bad` : `text-watch`)),
          children: t
        })]
      }, e))
    }), (0, $.jsxs)(`div`, {
      className: `grid gap-4 lg:grid-cols-[minmax(0,20rem)_minmax(0,1fr)]`,
      children: [(0, $.jsxs)(Jj, {
        className: `min-w-0`,
        children: [(0, $.jsx)(`h2`, {
          className: `font-display text-lg font-semibold`,
          children: `評價假設`
        }), (0, $.jsx)(`p`, {
          className: `mt-1 text-xs leading-relaxed text-muted`,
          children: `現價為 ${PRICE_DATE} 收盤 $${Y(VAL_DEFAULTS.price, 2)}。營收＝資金模型算力產能＋非算力服務。稅率 ${hA(VAL_DEFAULTS.tax * 100, 1)}（${TXQ.taxNote}；${TXQ.cashTaxNote}）。淨負債用 ${CALQ.valuationMD} 債務本金 ${Y(LATEST_Q.debtPrincipal, 2)} − 現金 ${Y(LATEST_Q.cash + LATEST_Q.marketable, 2)}，與期初現金同日。`
        }), (0, $.jsxs)(`div`, {
          className: `mt-4 space-y-3`,
          children: [(0, $.jsx)(kM, {
            label: `現價 US$`,
            hint: `${Y(n.price, 2)}（${mdQ(PRICE_DATE)} 收盤）`,
            children: (0, $.jsx)(OM, {
              value: n.price,
              onChange: e => i({
                price: e
              }),
              step: .01
            })
          }), (0, $.jsx)(kM, {
            label: `評價股數（十億，不含可轉債轉股）`,
            hint: `季末流通＋${TXQ.sharesNote} · 含可轉債／強制轉換特別股轉股 ${Y(h,4)}`,
            children: (0, $.jsx)(OM, {
              value: n.shares,
              onChange: e => i({
                shares: e
              }),
              step: .001
            })
          }), (0, $.jsx)(kM, {
            label: `淨負債 US$bn`,
            hint: `${TXQ.netDebtNote} · 評價淨負債 ${Y(r.v.netDebt, 2)}（含債務處理可轉債與持股調整）`,
            children: (0, $.jsx)(OM, {
              value: n.netDebt,
              onChange: e => i({
                netDebt: e
              })
            })
          }), (0, $.jsx)(kM, {
            label: `WACC`,
            hint: hA(n.wacc * 100, 1),
            children: (0, $.jsx)(`input`, {
              type: `range`,
              min: 80,
              max: 150,
              value: Math.round(n.wacc * 1e3),
              onChange: e => i({
                wacc: Number(e.target.value) / 1e3
              }),
              className: `w-full accent-accent`
            })
          }), (0, $.jsx)(kM, {
            label: `永續成長 g`,
            hint: hA(n.g * 100, 1),
            children: (0, $.jsx)(`input`, {
              type: `range`,
              min: 15,
              max: 45,
              value: Math.round(n.g * 1e3),
              onChange: e => i({
                g: Number(e.target.value) / 1e3
              }),
              className: `w-full accent-accent`
            })
          }), (0, $.jsx)(kM, {
            label: `終值維持性 CapEx（占 D&A）`,
            hint: hA(n.maintRatio * 100, 0),
            children: (0, $.jsx)(`input`, {
              type: `range`,
              min: 40,
              max: 120,
              value: Math.round(n.maintRatio * 100),
              onChange: e => i({
                maintRatio: Number(e.target.value) / 100
              }),
              className: `w-full accent-accent`
            })
          }), (0, $.jsx)(kM, {
            label: `EV/EBITDA 倍數`,
            hint: `${Y(n.evEbitda,1)}x`,
            children: (0, $.jsx)(OM, {
              value: n.evEbitda,
              onChange: e => i({
                evEbitda: e
              }),
              step: .5
            })
          }), (0, $.jsx)(kM, {
            label: `EV/EBITDA 錨定年度`,
            hint: `${PERIOD_LABELS[n.evYear ?? 1]}${(n.evYear ?? 1) >= CALQ.evDiscFrom ? `（以 WACC 折回 ${CALQ.targetText}）` : ``}`,
            children: (0, $.jsx)(`div`, {
              className: `grid grid-cols-4 gap-1 rounded-lg bg-surface p-1`,
              children: [1, 2, 3, 4].map(e => (0, $.jsx)(`button`, {
                onClick: () => i({
                  evYear: e
                }),
                className: Vj(`h-9 rounded-md text-sm font-medium`, (n.evYear ?? 1) === e ? `bg-ink text-accent-fg` : `text-muted hover:bg-card`),
                children: PERIOD_LABELS[e]
              }, e))
            })
          }), (0, $.jsx)(kM, {
            label: `DCF 股權為負時的處理`,
            hint: n.dcfMode === `option` ? `選擇權（Merton）` : `0 截斷`,
            children: (0, $.jsx)(`div`, {
              className: `grid grid-cols-2 gap-1 rounded-lg bg-surface p-1`,
              children: [[`zero`, `0 截斷`], [`option`, `選擇權`]].map(([e, t]) => (0, $.jsx)(`button`, {
                onClick: () => i({
                  dcfMode: e
                }),
                className: Vj(`h-9 rounded-md text-sm font-medium`, n.dcfMode === e ? `bg-ink text-accent-fg` : `text-muted hover:bg-card`),
                children: t
              }, e))
            })
          }), n.dcfMode === `option` && (0, $.jsx)(kM, {
            label: `企業價值波動率 σ`,
            hint: `${hA(n.sigma*100,0)} · 無風險利率 ${hA(n.rf*100,1)}、期間 ${n.optT} 年`,
            children: (0, $.jsx)(OM, {
              value: n.sigma * 100,
              onChange: e => i({
                sigma: Math.max(5, e) / 100
              }),
              step: 5
            })
          }), (0, $.jsxs)(`p`, {
            className: `text-xs text-muted`,
            children: [`DCF 失效只在 WACC ≤ g 或常態化 FCF ≤ 0 時成立，此時自動排除、權重改為 EV/EBITDA 100%。`, (0, $.jsx)(tipQ, {
              t: `0 截斷：股權價值＝MAX(0, 企業價值 − 淨負債 ＋ 新股現值)÷融資後股數，反映有限責任，但忽略翻身的可能。選擇權（Merton）：把股權視為以企業價值為標的、淨負債為履約價的買權（Black-Scholes，期間 ${Y(n.optT, 2)} 年、無風險利率 ${hA(n.rf * 100, 2)}），在 0 附近連續，並保留時間價值；代價是多一個波動率假設。注意選擇權法同樣會提高股權為正的情境（例如保守情境），因為它包含時間價值。`,
              w: 440
            })]
          }), (0, $.jsx)(`p`, {
            className: `text-xs leading-relaxed text-muted`,
            children: `${PERIODS[0]} 營收對照公司指引 ${REV_GUIDE_TXT}（年初至今 ${Y(ACTUAL_1H.revenue, 3)} 已入帳）；營收由 MW × 每 MW 年收入驅動，不回推指引。${PERIODS[1]} 無公司指引；市場共識 ${PERIODS[1]} 營收 ${Y(CONSENSUS.annualEstimates[PERIODS[1]].revenue, 2)}bn（${CONSENSUS.annualEstimates.source}，擷取 ${CONSENSUS.annualEstimates.retrieved}，${CONSENSUS.annualEstimates.tag}）僅供對照，見「市場共識」分頁。改資金模型的 MW／CapEx／預付，此頁營收、UFCF、目標價與結論會跟著動。DCF 折現期數按評價日至各期期末（${CALQ.tEnd.map(x => Y(x, 1)).join('／')} 年）。`
          })]
        })]
      }), (0, $.jsxs)(`div`, {
        className: `min-w-0 space-y-4`,
        children: [(0, $.jsx)(navQ, {
          groups: VMODS,
          cur: a,
          set: o
        }), a === 0 && (0, $.jsxs)(Jj, {
          className: `max-w-full overflow-x-auto`,
          children: [(0, $.jsx)(`h2`, {
            className: `font-display text-lg font-semibold`,
            children: `損益簡表（類損益表：歷史 ${HIST_PL[0].year}–${HIST_PL[HIST_PL.length - 1].year}＋前瞻五期）` // v4.5：兩端讀 historicalPL（最後一格＝年初至今實際，滾動時隨資料更新）
          }), (0, $.jsx)(`p`, {
            className: `mt-1 text-xs leading-relaxed text-muted`,
            children: `前瞻算力收入＝平均在役 MW × 每 MW 年收入（Tokenomics 正向推導）；${TXQ.otherRevNote}。${PERIODS[0]} 欄＝年初至今實際（季報）＋模型期，可直接對照公司全年指引 ${REV_GUIDE_TXT}；${PERIODS[1]} 起為純模型。DCF 只折現評價日之後的現金流。`
          }), (0, $.jsx)(`div`, {
            className: `mt-3`,
            children: (0, $.jsx)(BM, {
              cols: [...HIST_PL.map((e, t) => t === HIST_PL.length - 1 ? `${e.year} 實際` : e.year), ...PERIODS.map(p => `${p}E`)], // v0.1b：欄名讀 historicalPL 與日曆
              rows: [
                [`營收`, null],
                [`${HIST_PL[3].year} 實際營收（已實現）`, [NaN, NaN, NaN, NaN, HIST_PL[3].revenue, 0, 0, 0, 0], void 0, void 0, void 0, `${PERIODS[0]}E 欄＝${HIST_PL[3].year} 實際＋${CALQ.stubWord}模型。`],
                [`算力收入（模型期）`, [NaN, NaN, NaN, NaN, ...s.map(e => e.gpu)], void 0, void 0, void 0, `＝期初 RPO 轉換（產能約束後）＋新簽約。${PERIODS[0]}E 欄只含${CALQ.stubWord}。`],
                [`　其中：期初 RPO 轉換`, [NaN, NaN, NaN, NaN, ...s.map(e => e.fundingRev)]],
                [`　其中：新簽約`, [NaN, NaN, NaN, NaN, ...s.map(e => e.inYear)], void 0, void 0, void 0, `＝(容量 − 期初 RPO 排程)×新產能簽約率。`],
                [`非算力服務（模型期）`, [NaN, NaN, NaN, NaN, ...s.map(e => e.services)]],
                [`總營收`, [...HIST_PL.map(e => e.revenue), ...s.map(e => e.fyRevenue)], void 0, void 0, `tot`, `${PERIODS[0]}E 對照公司全年指引 ${REV_GUIDE_TXT}。`],
                [`營收 YoY`, [NaN, HIST_PL[1].revenue / HIST_PL[0].revenue - 1, HIST_PL[2].revenue / HIST_PL[1].revenue - 1, NaN, ...s.map((e, t) => t === 0 ? e.fyRevenue / HIST_PL[2].revenue - 1 : e.fyRevenue / s[t - 1].fyRevenue - 1)].map(e => e * 100), void 0, void 0, void 0, void 0, `%`],
                [`獲利`, null],
                [`GAAP 營業利益`, [...HIST_PL.map(e => e.opInc), ...s.map(e => e.fyOpInc)], void 0, void 0, void 0, void 0, void 0, 2],
                [`營利率`, [...HIST_PL.map(e => e.opInc / e.revenue), ...s.map(e => e.fyOpInc / e.fyRevenue)].map(e => e * 100), void 0, void 0, void 0, void 0, `%`],
                [`D&A`, [NaN, NaN, NaN, ACTUAL_1H.da, ...s.map(e => e.fyDa)], void 0, void 0, void 0, `年初至今實際 ${Y(ACTUAL_1H.da, 3)}（季報）。`],
                [`EBITDA（EBIT＋D&A）`, [NaN, NaN, NaN, HIST_PL[3].opInc + ACTUAL_1H.da, ...s.map(e => e.fyEbitda)], void 0, void 0, `tot`],
                [`利息（含瀑布新債，模型期）`, [NaN, NaN, NaN, ACTUAL_1H.interest, ...s.map((e, t) => t === 0 ? ACTUAL_1H.interest + e.interest : e.interest)], void 0, void 0, void 0, `${PERIOD_FY[0]}E＝年初至今實際 ${Y(ACTUAL_1H.interest, 3)}＋剩餘期間模型；之後含瀑布新債與可轉債票息。`],
                [`所得稅（NOL 後）`, [NaN, NaN, NaN, NaN, ...s.map(e => e.tax)], void 0, void 0, void 0, `虧損年不認列稅盾；獲利年 ${hA(n.tax * 100, 1)}、NOL 抵扣上限 ${hA(n.nolUsePct * 100, 0)}。`],
                [`淨利`, [...HIST_PL.map(e => e.ni), ...s.map(e => e.fyNi)], void 0, void 0, `tot`, void 0, void 0, 2],
                [`淨利率`, [...HIST_PL.map(e => e.ni / e.revenue), ...s.map(e => e.fyNi / e.fyRevenue)].map(e => e * 100), void 0, void 0, void 0, void 0, `%`],
                [`每股`, null],
                [`股數（含 ATM 上限，每年 +1% SBC 稀釋）`, [NaN, NaN, NaN, NaN, ...s.map(e => e.shares)], void 0, void 0, void 0, `SBC 約 ${Y(n.sbc, 2)}bn／年 ÷ 現價 $${Y(n.price, 0)} ≈ ${Y(n.sbc / n.price * 1e3, 1)}m 股（約 ${hA(n.sbc / n.price / n.shares * 100, 1)}／年），模板固定 1%（偏保守）。起點含價內可轉債轉股。`, `bn 股`, 4],
                [`GAAP EPS`, [...HIST_PL.map(e => e.eps), ...s.map(e => e.fyEps)], void 0, void 0, void 0, `${PERIOD_FY[0]}E＝年初至今實際 EPS ${Y(HIST_PL[3].eps, 2)} ＋ 剩餘期間淨利 ÷ 股數。`, `US$`, 2],
                [`EPS（加回 SBC）`, [...HIST_PL.map(e => e.ngEps), ...s.map(e => e.fyNgEps)], void 0, void 0, void 0, `＝(淨利＋SBC)÷股數；SBC 全額加回、不做稅盾調整（與公司非 GAAP EPS 口徑不同）。`, `US$`, 2]
              ]
            })
          }), (0, $.jsxs)(`p`, {
            className: `mt-3 text-xs text-muted`,
            children: [`${TXQ.guideLine}。本表 ${PERIODS[0]}E 總營收 `, Y(HIST_PL[3].revenue + s[0].revenue, 2), `bn（年初至今 ${Y(HIST_PL[3].revenue, 2)} 實際）。${PERIODS[1]} 營收 `, Y(s[1].revenue, 1), `bn。稅率 ${hA(n.tax * 100, 1)}（${TXQ.taxNote}）；期初虧損扣抵 ${Y(n.nol,1)}bn，抵扣上限為應稅所得 ${hA(n.nolUsePct * 100, 0)}。DCF 的現金稅同步使用 NOL。`]
          }), (0, $.jsx)(`div`, {
            className: `mt-3 overflow-x-auto`,
            children: (0, $.jsxs)(`table`, {
              className: `w-full min-w-[640px] text-xs`,
              children: [(0, $.jsx)(`thead`, {
                children: (0, $.jsxs)(`tr`, {
                  className: `text-muted`,
                  children: [(0, $.jsx)(`th`, {
                    className: `py-2 text-left font-medium`,
                    children: `前瞻假設（EBITDA 率與服務營收在資金模型編輯）`
                  }), PERIOD_LABELS.map(e => (0, $.jsx)(`th`, {
                    className: `py-2 text-right font-medium`,
                    children: e
                  }, e))]
                })
              }), (0, $.jsxs)(`tbody`, {
                children: [(0, $.jsxs)(`tr`, {
                  children: [(0, $.jsxs)(`td`, {
                    className: `font-medium`,
                    children: [`EBITDA 率（連動資金模型）`, (0, $.jsx)(tipQ, {
                      t: `由資金模型左欄的「起始 EBITDA 率」與「穩態 EBITDA 率」線性推得（${PERIODS[0]} 起始 → ${PERIODS[4]} 穩態），資金端的EBITDAR 率使用同一組數字。要修改請到資金模型左欄。EBITDA 已扣營業租賃成本；營業利益＝EBITDA − 車隊 D&A。`
                    })]
                  }), e.years.map((e, t) => (0, $.jsx)(`td`, {
                    className: `p-1 text-right font-mono tabular-nums`,
                    children: hA(e.ebM * 100, 1)
                  }, t))]
                }), (0, $.jsxs)(`tr`, {
                  children: [(0, $.jsxs)(`td`, {
                    className: `font-medium`,
                    children: [`D&A（車隊折舊，計算）`, (0, $.jsx)(tipQ, {
                      t: `＝(期初毛 PP&E ＋ 本期成長型 CapEx×½) ÷ GPU 經濟壽命 × 期間長度。期初 PP&E 基礎 ${Y(DEFAULTS.ppeOpen, 1)}bn（${TXQ.ppeOpenNote}；評價日 PP&E 淨額 ${Y(LATEST_Q.ppe, 1)}、在建 ${Y(LATEST_Q.cip, 1)}）；壽命 ${DEFAULTS.gpuLife} 年。汰換型 CapEx 取代已折舊完的舊設備，不增加折舊基礎。`,
                      w: 440
                    })]
                  }), s.map((e, t) => (0, $.jsx)(`td`, {
                    className: `p-1 text-right font-mono tabular-nums`,
                    children: Y(e.da, 1)
                  }, t))]
                }), (0, $.jsxs)(`tr`, {
                  children: [(0, $.jsx)(`td`, {
                    className: `font-medium`,
                    children: `非算力服務營收（連動資金模型）`
                  }), e.years.map((e, t) => (0, $.jsx)(`td`, {
                    className: `p-1 text-right font-mono tabular-nums`,
                    children: Y(e.servicesRev, 1)
                  }, t))]
                })]
              })]
            })
          })]
        }), a === 1 && (0, $.jsxs)(Jj, {
          className: `max-w-full overflow-x-auto`,
          children: [(0, $.jsx)(hdrQ, {
            title: `可比公司（以 EV/Sales 為主）`,
            tip: COMPANY_DATA.peers.textHtml.headerTip
          }), (0, $.jsx)(`div`, {
            className: `mt-3`,
            style: xstyQ.wrap,
            children: (0, $.jsxs)(`table`, {
              style: {
                ...xstyQ.table,
                minWidth: 1080
              },
              children: [(0, $.jsx)(`thead`, {
                children: (0, $.jsx)(`tr`, {
                  children: [`公司`, `市值`, `淨負債`, `EV`, `TTM 營收`, `EV/Sales（TTM）`, `含租賃 EV/Sales`, `TTM EBITDA`, `EV/EBITDA`, `資料期`, `備註`].map((e, t) => (0, $.jsx)(`th`, {
                    style: t === 0 ? xstyQ.thL : t === 10 ? {
                      ...xstyQ.th,
                      textAlign: `left`
                    } : t === 5 ? {
                      ...xstyQ.th,
                      background: `#0f5c61`
                    } : xstyQ.th,
                    children: e
                  }, e))
                })
              }), (0, $.jsxs)(`tbody`, {
                children: [...lM.map((e, t) => (0, $.jsxs)(`tr`, {
                  children: [(0, $.jsxs)(`td`, {
                    style: {
                      ...xstyQ.tdL,
                      background: t % 2 ? `#fbfaf7` : `#fff`
                    },
                    children: [(0, $.jsx)(`b`, {
                      children: e.labelHtml
                    }), ` `, (0, $.jsx)(`span`, {
                      style: {
                        color: `#6b7280`,
                        fontSize: 11
                      },
                      children: e.roleHtml
                    })]
                  }), [e.mkt, e.netDebt, e.ev, e.rev].map((e, n) => (0, $.jsx)(`td`, {
                    style: {
                      ...xstyQ.td,
                      background: t % 2 ? `#fbfaf7` : `#fff`
                    },
                    children: nfmtQ(e, n === 3 ? 2 : 1)
                  }, n)), (0, $.jsx)(`td`, {
                    style: {
                      ...xstyQ.td,
                      background: `#e6f1f1`,
                      fontWeight: 700
                    },
                    children: `${Y(e.ev / e.rev, 1)}x`
                  }), (0, $.jsx)(`td`, {
                    style: {
                      ...xstyQ.td,
                      background: t % 2 ? `#fbfaf7` : `#fff`
                    },
                    children: `${Y((e.ev + (e.opl || 0)) / e.rev, 1)}x`
                  }), (0, $.jsx)(`td`, {
                    style: {
                      ...xstyQ.td,
                      background: t % 2 ? `#fbfaf7` : `#fff`,
                      color: e.ebitda < 0 ? `#9f1239` : `#1f2937`
                    },
                    children: nfmtQ(e.ebitda, 2)
                  }), (0, $.jsx)(`td`, {
                    style: {
                      ...xstyQ.td,
                      background: t % 2 ? `#fbfaf7` : `#fff`,
                      color: `#6b7280`
                    },
                    children: e.ebitda > 0 ? `${Y(e.ev / e.ebitda, 1)}x` : `n/m`
                  }), (0, $.jsx)(`td`, {
                    style: {
                      ...xstyQ.td,
                      background: t % 2 ? `#fbfaf7` : `#fff`,
                      fontSize: 10.5,
                      color: `#6b7280`
                    },
                    children: e.asOf
                  }), (0, $.jsx)(`td`, {
                    style: {
                      ...xstyQ.td,
                      textAlign: `left`,
                      whiteSpace: `normal`,
                      minWidth: 260,
                      fontSize: 11,
                      color: `#6b7280`,
                      fontFamily: `ui-sans-serif, system-ui, sans-serif`,
                      background: t % 2 ? `#fbfaf7` : `#fff`
                    },
                    children: e.noteHtml
                  })]
                }, e.ticker)), (0, $.jsxs)(`tr`, {
                  children: [(0, $.jsx)(`td`, {
                    style: {
                      ...xstyQ.tdL,
                      background: `#f3f4f6`,
                      fontWeight: 700
                    },
                    children: `同業中位數`
                  }), [NaN, NaN, NaN, NaN].map((e, t) => (0, $.jsx)(`td`, {
                    style: {
                      ...xstyQ.td,
                      background: `#f3f4f6`
                    },
                    children: ``
                  }, t)), (0, $.jsx)(`td`, {
                    style: {
                      ...xstyQ.td,
                      background: `#e6f1f1`,
                      fontWeight: 700
                    },
                    children: `${Y(wM(lM.map(e => e.ev / e.rev)), 1)}x`
                  }), (0, $.jsx)(`td`, {
                    style: {
                      ...xstyQ.td,
                      background: `#f3f4f6`
                    },
                    children: `${Y(wM(lM.map(e => (e.ev + (e.opl || 0)) / e.rev)), 1)}x`
                  }), [0, 1, 2, 3].map(e => (0, $.jsx)(`td`, {
                    style: {
                      ...xstyQ.td,
                      background: `#f3f4f6`
                    },
                    children: ``
                  }, `m${e}`))]
                }), (0, $.jsxs)(`tr`, {
                  children: [(0, $.jsx)(`td`, {
                    style: {
                      ...xstyQ.tdL,
                      background: `#eef2f7`,
                      fontWeight: 700
                    },
                    children: `${COMPANY_DATA.meta.ticker}（TTM）`
                  }), [CALL_FACTS.mktCapLast, NDT, CALL_FACTS.mktCapLast + NDT, CALL_FACTS.ttmRev].map((e, t) => (0, $.jsx)(`td`, {
                    style: {
                      ...xstyQ.td,
                      background: `#eef2f7`,
                      fontWeight: 700
                    },
                    children: nfmtQ(e, t === 3 ? 2 : 1)
                  }, t)), (0, $.jsx)(`td`, {
                    style: {
                      ...xstyQ.td,
                      background: `#e6f1f1`,
                      fontWeight: 700
                    },
                    children: `${Y((CALL_FACTS.mktCapLast + NDT) / CALL_FACTS.ttmRev, 1)}x`
                  }), (0, $.jsx)(`td`, {
                    style: {
                      ...xstyQ.td,
                      background: `#eef2f7`,
                      fontWeight: 700
                    },
                    children: `${Y((CALL_FACTS.mktCapLast + NDT + LATEST_Q.opLeaseLiab) / CALL_FACTS.ttmRev, 1)}x`
                  }), (0, $.jsx)(`td`, {
                    style: {
                      ...xstyQ.td,
                      background: `#eef2f7`
                    },
                    children: nfmtQ(CALL_FACTS.ttmOpInc + CALL_FACTS.ttmDa, 2)
                  }), (0, $.jsx)(`td`, {
                    style: {
                      ...xstyQ.td,
                      background: `#eef2f7`
                    },
                    children: `${Y((CALL_FACTS.mktCapLast + NDT) / (CALL_FACTS.ttmOpInc + CALL_FACTS.ttmDa), 1)}x`
                  }), (0, $.jsx)(`td`, {
                    style: {
                      ...xstyQ.td,
                      background: `#eef2f7`,
                      fontSize: 10.5
                    },
                    children: `TTM 至 ${CALQ.filedQLabel}`
                  }), (0, $.jsx)(`td`, {
                    style: {
                      ...xstyQ.td,
                      textAlign: `left`,
                      whiteSpace: `normal`,
                      fontSize: 11,
                      background: `#eef2f7`,
                      fontFamily: `ui-sans-serif, system-ui, sans-serif`
                    },
                    children: `EV 用 ${mdQ(CALL_FACTS.priceDate)} 市值與 ${CALQ.valuationMD} 淨負債 ${Y(NDT, 1)}；含租賃再加營業租賃負債 ${Y(LATEST_Q.opLeaseLiab, 1)}。TTM GAAP 營業損益 ${Y(CALL_FACTS.ttmOpInc, 3)}＋D&A ${Y(CALL_FACTS.ttmDa, 3)}`
                  })]
                }), (0, $.jsxs)(`tr`, {
                  children: [(0, $.jsx)(`td`, {
                    style: {
                      ...xstyQ.tdL,
                      background: `#fff`
                    },
                    children: `${COMPANY_DATA.meta.ticker}（模型 ${PERIODS[1]}E）`
                  }), [y, r.v.netDebt, b, x].map((e, t) => (0, $.jsx)(`td`, {
                    style: {
                      ...xstyQ.td,
                      background: `#fff`
                    },
                    children: nfmtQ(e, t === 3 ? 2 : 1)
                  }, t)), (0, $.jsx)(`td`, {
                    style: {
                      ...xstyQ.td,
                      background: `#e6f1f1`
                    },
                    children: `${Y(b / x, 1)}x`
                  }), (0, $.jsx)(`td`, {
                    style: {
                      ...xstyQ.td,
                      background: `#fff`
                    },
                    children: `${Y((b + LATEST_Q.opLeaseLiab) / x, 1)}x`
                  }), (0, $.jsx)(`td`, {
                    style: {
                      ...xstyQ.td,
                      background: `#fff`
                    },
                    children: nfmtQ(S, 2)
                  }), (0, $.jsx)(`td`, {
                    style: {
                      ...xstyQ.td,
                      background: `#fff`
                    },
                    children: `${Y(C, 1)}x`
                  }), (0, $.jsx)(`td`, {
                    style: {
                      ...xstyQ.td,
                      background: `#fff`,
                      fontSize: 10.5
                    },
                    children: `模型`
                  }), (0, $.jsx)(`td`, {
                    style: {
                      ...xstyQ.td,
                      textAlign: `left`,
                      whiteSpace: `normal`,
                      fontSize: 11,
                      background: `#fff`,
                      fontFamily: `ui-sans-serif, system-ui, sans-serif`
                    },
                    children: `市值用模型股數 ${Y(h,3)}bn（含轉股）；分母為模型 ${PERIODS[1]} 營收`
                  })]
                })]
              })]
            })
          }), (0, $.jsxs)(`p`, {
            className: `mt-3 text-xs leading-relaxed text-muted`,
            children: [`讀法：${COMPANY_DATA.meta.ticker} 的 TTM EV/Sales `, Y((CALL_FACTS.mktCapLast + NDT) / CALL_FACTS.ttmRev, 1), `x，約為同業中位數 `, Y(wM(lM.map(e => e.ev / e.rev)), 1), `x 的 `, Y((CALL_FACTS.mktCapLast + NDT) / CALL_FACTS.ttmRev / wM(lM.map(e => e.ev / e.rev)), 1), ` 倍。`, (0, $.jsx)(tipQ, {
              t: COMPANY_DATA.peers.textHtml.readingTip,
              w: 480
            })]
          }), (0, $.jsx)(`p`, {
            className: `mt-3 text-xs leading-relaxed text-muted`,
            children: COMPANY_DATA.peers.textHtml.caveat
          })]

        }), a === 2 && (0, $.jsxs)(Jj, {
          className: `max-w-full overflow-x-auto`,
          children: [(0, $.jsx)(`h2`, {
            className: `font-display text-lg font-semibold`,
            children: `DCF（與資金模型 Cash CapEx 連動）`
          }), (0, $.jsx)(`p`, {
            className: `mt-1 text-xs leading-relaxed text-muted`,
            children: `UFCF = EBIT×(1−t)＋D&A−Cash CapEx−營運資金。Cash CapEx 與利息取自資金模型。${VAL_DEFAULTS.tvBasis === `ufcf` ? `終值以末期（${PERIODS[4]}）UFCF 為基準：MW 觸頂後與終值年含穩態 GPU 汰換 CapEx（已連網 MW × 每 MW ÷ 壽命），GPU 不是永續資產；末期若仍有成長型 CapEx 則加回（成長由 g 表達）。` : `五期 UFCF 多為負（建置期），終值不用 FY30 UFCF，而用常態化 FCF＝EBIT×(1−t)＋D&A−維持性 CapEx（D&A×維持比率）。`}DCF 已內含建置支出，不再重複扣期末現金缺口。每股值以 0 為下限：股權價值為負代表企業價值低於淨負債，股東有限責任、股價下限為 0。`
          }), (0, $.jsxs)(`table`, {
            className: `mt-3 w-full min-w-[640px] text-xs`,
            children: [(0, $.jsx)(`thead`, {
              children: (0, $.jsxs)(`tr`, {
                className: `text-muted`,
                children: [(0, $.jsx)(`th`, {
                  className: `py-2 text-left font-medium`,
                  children: `US$bn`
                }), s.map(e => (0, $.jsx)(`th`, {
                  className: `py-2 text-right font-medium`,
                  children: e.year
                }, e.year))]
              })
            }), (0, $.jsx)(`tbody`, {
              children: [
                [`EBIT`, s.map(e => e.opInc)],
                [`利息（資金）`, s.map(e => e.interest)],
                [`D&A`, s.map(e => e.da)],
                [`Cash CapEx`, s.map(e => e.cashCapex)],
                [`UFCF`, s.map(e => e.ufcf)],
                [`PV`, c.pv]
              ].map(([e, t]) => (0, $.jsxs)(`tr`, {
                className: `border-t border-border`,
                children: [(0, $.jsx)(`td`, {
                  className: `py-1.5 font-medium`,
                  children: e
                }), t.map((t, n) => (0, $.jsx)(`td`, {
                  className: Vj(`py-1.5 text-right font-mono tabular-nums`, e === `UFCF` && t < 0 && `text-bad`),
                  children: Y(t, 1)
                }, n))]
              }, e))
            })]
          }), (0, $.jsxs)(`div`, {
            className: `mt-4 grid gap-3 sm:grid-cols-2 lg:grid-cols-5`,
            children: [(0, $.jsx)(jM, {
              k: `FCF 現值`,
              v: mA(c.pvFcf)
            }), (0, $.jsx)(jM, {
              k: VAL_DEFAULTS.tvBasis === `ufcf` ? `終值基準 FCF（${PERIODS[4]} UFCF）` : `常態化 FCF（${PERIODS[4]}）`,
              v: mA(c.normFcf)
            }), (0, $.jsx)(jM, {
              k: `終值現值`,
              v: mA(c.pvTv)
            }), (0, $.jsx)(jM, {
              k: `終值占比`,
              v: hA(c.tvShare * 100, 0)
            }), (0, $.jsx)(jM, {
              k: `新股募得現金（現值）`,
              v: mA(c.pvEquityRaised)
            }), (0, $.jsx)(jM, {
              k: `融資後股數（bn）`,
              v: Y(c.postShares, 3)
            }), (0, $.jsx)(jM, {
              k: `DCF 每股（${c.invalid ? `失效` : n.dcfMode === `option` ? `選擇權` : `0 截斷`}）`,
              v: c.invalid ? `失效：${c.invalidReason}` : `$${Y(c.perShare,1)}（原始 $${Y(c.perShareRaw,1)}；0 截斷 $${Y(c.zeroPerShare,1)}；選擇權 $${Y(c.optPerShare,1)}）；推到 ${CALQ.targetDate} $${Y(c.perShareT,1)}`
            })]
          }), (0, $.jsx)(`h3`, {
            className: `mt-5 text-sm font-medium`,
            children: `WACC × g 敏感度（每股 US$）`
          }), (0, $.jsxs)(`table`, {
            className: `mt-2 w-full min-w-[520px] text-xs`,
            children: [(0, $.jsx)(`thead`, {
              children: (0, $.jsxs)(`tr`, {
                className: `text-muted`,
                children: [(0, $.jsx)(`th`, {
                  className: `py-2 text-left font-medium`,
                  children: `g＼WACC`
                }), [`9.0%`, `10.0%`, `11.0%`, `12.0%`, `13.0%`].map(e => (0, $.jsx)(`th`, {
                  className: `py-2 text-right font-medium`,
                  children: e
                }, e))]
              })
            }), (0, $.jsx)(`tbody`, {
              children: [`2.0%`, `2.5%`, `3.0%`, `3.5%`, `4.0%`].map((e, t) => (0, $.jsxs)(`tr`, {
                className: `border-t border-border`,
                children: [(0, $.jsx)(`td`, {
                  className: `py-1.5 font-medium`,
                  children: e
                }), g[t].map((e, n) => (0, $.jsx)(`td`, {
                  className: Vj(`py-1.5 text-right font-mono tabular-nums`, t === 2 && n === 2 && `bg-accent-soft font-medium`),
                  children: Number.isFinite(e) ? Y(e, 0) : `—`
                }, n))]
              }, e))
            })]
          })]
        }), a === 4 && RV && (0, $.jsxs)(Jj, {
          className: `max-w-full overflow-x-auto space-y-3`,
          children: [(0, $.jsx)(hdrQ, {
            title: `反向 DCF：現價 $${Y(RV.price, 2)} 隱含什麼樣的收入與成本`,
            tip: `不是問「值多少」，而是問「要值現價，營運要長什麼樣」。做法：固定其他所有輸入，調整單一或組合變數，直到 DCF 每股（目前下限方式：${n.dcfMode === `option` ? `選擇權` : `0 截斷`}）＝ 現價，每一格都重跑完整的資金瀑布與評價。收入端以「每 MW 年收入倍數」整體調整五期路徑；成本端分為資本成本（每 MW 建置成本倍數，同時影響 CapEx、汰換與車隊折舊）與營運成本（穩態 EBITDA 率）。`,
            w: 500
          }), (0, $.jsx)(`div`, {
            style: xstyQ.wrap,
            children: (0, $.jsxs)(`table`, {
              style: {
                ...xstyQ.table,
                minWidth: 760
              },
              children: [(0, $.jsx)(`thead`, {
                children: (0, $.jsx)(`tr`, {
                  children: [`單一槓桿（其他不變）`, `目前`, `市價隱含`, `變動`, `對照`].map((e, t) => (0, $.jsx)(`th`, {
                    style: t === 0 ? xstyQ.thL : t === 4 ? {
                      ...xstyQ.th,
                      textAlign: `left`
                    } : xstyQ.th,
                    children: e
                  }, e))
                })
              }), (0, $.jsx)(`tbody`, {
                children: [
                  [`${PERIODS[4]} 每 MW 年收入（US$m）`, Y(RV.rev30, 1), Number.isFinite(RV.R) ? Y(RV.rev30 * RV.R, 1) : `不可達`, Number.isFinite(RV.R) ? hA((RV.R - 1) * 100, 0) : `—`, TXQ.rvRevNote],
                  [`每 MW 建置成本（US$m）`, Y(RV.cost30, 1), Number.isFinite(RV.C) ? Y(RV.cost30 * RV.C, 1) : `不可達`, Number.isFinite(RV.C) ? hA((RV.C - 1) * 100, 0) : `—`, TXQ.rvCostNote],
                  [`穩態 EBITDA 率（${PERIODS[4]}）`, hA(RV.eb30 * 100, 1), Number.isFinite(RV.Eb) ? hA(RV.Eb * 100, 1) : `不可達（>99%）`, Number.isFinite(RV.Eb) ? `${Y((RV.Eb - RV.eb30) * 100, 1)} pt` : `—`, `AI 雲端（算力）EBITDA 率，非 AI 事業不動；可觀察 neocloud 區間：IREN 約 35%、CRWV 約 59%`],
                  [`（對照）加權目標價＝現價所需每 MW 年收入`, Y(RV.rev30, 1), Number.isFinite(RV.Rt) ? Y(RV.rev30 * RV.Rt, 1) : `不可達`, Number.isFinite(RV.Rt) ? hA((RV.Rt - 1) * 100, 0) : `—`, `含 EV/EBITDA ${Y(n.evEbitda,1)}x（${PERIOD_LABELS[n.evYear ?? 1]}）腿；非純反向 DCF`]
                ].map((e, t) => (0, $.jsx)(`tr`, {
                  children: e.map((e, n) => (0, $.jsx)(`td`, {
                    style: {
                      ...(n === 0 ? xstyQ.tdL : xstyQ.td),
                      background: t % 2 ? `#fbfaf7` : `#fff`,
                      fontWeight: n === 2 ? 700 : 400,
                      ...(n === 4 ? {
                        textAlign: `left`,
                        whiteSpace: `normal`,
                        fontSize: 11,
                        color: `#6b7280`,
                        fontFamily: `ui-sans-serif, system-ui, sans-serif`,
                        minWidth: 260
                      } : {})
                    },
                    children: e
                  }, n))
                }, t))
              })]
            })
          }), (0, $.jsx)(hdrQ, {
            title: `收入與成本的綜合影響：要值現價，${PERIODS[4]} 每 MW 年收入需要多少（US$m）`,
            tip: `列＝每 MW 建置成本（相對目前 $${Y(RV.cost30,1)}m 的倍數），欄＝穩態 EBITDA 率。每一格解出 DCF＝現價所需的每 MW 年收入（括號為相對目前 $${Y(RV.rev30,1)}m 的變動）。左上往右下的對角線，就是「收入 × 成本」的等價值線：成本愈低、利潤率愈高，所需收入愈低。黃底為目前設定。`,
            w: 480
          }), (0, $.jsx)(`div`, {
            style: xstyQ.wrap,
            children: (0, $.jsxs)(`table`, {
              style: {
                ...xstyQ.table,
                minWidth: 640
              },
              children: [(0, $.jsx)(`thead`, {
                children: (0, $.jsx)(`tr`, {
                  children: [`建置成本 ＼ 穩態 EBITDA 率`, ...RV.ebs.map(e => hA(e * 100, 0))].map((e, t) => (0, $.jsx)(`th`, {
                    style: t === 0 ? xstyQ.thL : xstyQ.th,
                    children: e
                  }, e))
                })
              }), (0, $.jsx)(`tbody`, {
                children: RV.caps.map((c, i) => (0, $.jsxs)(`tr`, {
                  children: [(0, $.jsx)(`td`, {
                    style: {
                      ...xstyQ.tdL,
                      background: i % 2 ? `#fbfaf7` : `#fff`
                    },
                    children: `$${Y(RV.cost30 * c, 1)}m（${c === 1 ? `目前` : (c > 1 ? `+` : ``) + hA((c - 1) * 100, 0)}）`
                  }), ...RV.grid[i].map((x, j) => {
                    let on = c === 1 && Math.abs(RV.ebs[j] - RV.eb30) < 1e-6;
                    return (0, $.jsx)(`td`, {
                      style: {
                        ...xstyQ.td,
                        background: on ? `#fff2a8` : i % 2 ? `#fbfaf7` : `#fff`,
                        fontWeight: on ? 700 : 400,
                        color: Number.isFinite(x) && x <= 1 ? `#0f6b4c` : `#1f2937`
                      },
                      children: Number.isFinite(x) ? `${Y(RV.rev30 * x, 1)}（${x >= 1 ? `+` : ``}${hA((x - 1) * 100, 0)}）` : `—`
                    }, j)
                  })]
                }, c))
              })]
            })
          }), (0, $.jsxs)(`p`, {
            className: `text-xs leading-relaxed text-muted`,
            children: [`讀法：以目前設定，現價要成立，需要 ${PERIODS[4]} 每 MW 年收入由 $${Y(RV.rev30,1)}m 升到 $${Number.isFinite(RV.R) ? Y(RV.rev30*RV.R,1) : `—`}m（約 +${Number.isFinite(RV.R) ? hA((RV.R-1)*100,0) : `—`}）；或建置成本降到 $${Number.isFinite(RV.C) ? Y(RV.cost30*RV.C,1) : `—`}m；單靠營運效率則需 EBITDA 率 ${Number.isFinite(RV.Eb) ? hA(RV.Eb*100,0) : `>99%`}，高於由下而上的上緣。綠字格代表不需提高單價即可支撐現價的組合。`, (0, $.jsx)(tipQ, {
              t: `單價與成本的組合比任何單一變數都更貼近現實：${TXQ.priceNote}；若新世代 GPU 同時推高每 MW 成本，兩者會互相抵銷。追蹤指標：新簽約的每 MW 年收入（法說與 8-K 揭露的合約金額 ÷ MW）、每 MW CapEx（CapEx ÷ 新增主動電力）、調整後營業利益率。`,
              w: 440
            })]
          })]
        }), a === 5 && (0, $.jsx)(ConsTabQ, { d: e, p: r, o: n, tr: TR, st: t }), a === 3 && (0, $.jsxs)(Jj, {
          children: [(0, $.jsxs)(`div`, {
            className: `flex flex-wrap items-center gap-2`,
            children: [(0, $.jsx)(`h2`, {
              className: `font-display text-lg font-semibold`,
              children: `目標價與區間`
            }), (0, $.jsx)(Uj, {
              tone: D,
              children: m.call
            }), (0, $.jsx)(Uj, {
              tone: m.funded ? `ok` : `bad`,
              children: m.funded ? `無需新股` : `需股權 ${Y(p,0)}bn`
            })]
          }), (0, $.jsxs)(`p`, {
            className: `mt-2 text-sm leading-relaxed`,
            children: [`加權目標價 `, (0, $.jsxs)(`span`, {
              className: `font-semibold tabular-nums`,
              children: [`$`, Y(m.blended, 0)]
            }), c.invalid ? `＝ EV/EBITDA（融資後）$${Y(f, 0)} × 100%（DCF 失效已排除）` : `＝ DCF $${Y(c.perShareT, 0)}（${n.dcfMode === `option` ? `選擇權` : `0 截斷`}，推到 ${CALQ.targetDate}）× ${hA(m.weights.dcf * 100, 0)} ＋ EV/EBITDA（融資後）$${Y(f, 0)} × ${hA(m.weights.pe * 100, 0)}`, `。相對現價 $`, Y(n.price, 2), ` 為`, ` `, (0, $.jsx)(`span`, {
              className: m.upside >= 0 ? `text-ok` : `text-bad`,
              children: hA(m.upside * 100, 0)
            }), `。五期需股權 `, Y(p, 0), `bn。`]
          }), TR.note ? (0, $.jsx)(`p`, {
            className: `mt-2 text-xs leading-relaxed text-muted`,
            children: TR.note
          }) : null, (0, $.jsx)(`p`, {
            className: `mt-2 text-sm leading-relaxed`,
            children: `${TR.head}；${TR.bLabel} $${Y(TR.B[0], 1)}–$${Y(TR.B[1], 1)}，${TR.pos}。${TR.judge}`
          }), (0, $.jsx)(`div`, {
            className: `mt-4 overflow-x-auto`,
            children: (0, $.jsxs)(`table`, {
              className: `w-full min-w-[640px] text-xs`,
              children: [(0, $.jsx)(`thead`, {
                children: (0, $.jsxs)(`tr`, {
                  className: `text-muted`,
                  children: [(0, $.jsx)(`th`, {
                    className: `py-2 text-left font-medium`,
                    children: `方法`
                  }), (0, $.jsx)(`th`, {
                    className: `py-2 text-left font-medium`,
                    children: `CapEx`
                  }), (0, $.jsx)(`th`, {
                    className: `py-2 text-left font-medium`,
                    children: `缺口處理`
                  }), (0, $.jsx)(`th`, {
                    className: `py-2 text-left font-medium`,
                    children: `股數`
                  }), (0, $.jsx)(`th`, {
                    className: `py-2 text-left font-medium`,
                    children: `淨負債`
                  })]
                })
              }), (0, $.jsx)(`tbody`, {
                children: bM.map(e => (0, $.jsxs)(`tr`, {
                  className: `border-t border-border`,
                  children: [(0, $.jsx)(`td`, {
                    className: `py-1.5 font-medium`,
                    children: e.method
                  }), (0, $.jsx)(`td`, {
                    className: `py-1.5 text-muted`,
                    children: e.capex
                  }), (0, $.jsx)(`td`, {
                    className: `py-1.5 text-muted`,
                    children: e.hole
                  }), (0, $.jsx)(`td`, {
                    className: `py-1.5 text-muted`,
                    children: e.shares
                  }), (0, $.jsx)(`td`, {
                    className: `py-1.5 text-muted`,
                    children: e.netDebt
                  })]
                }, e.method))
              })]
            })
          }), (0, $.jsx)(hdrQ, {
            title: `敏感度矩陣：穩態 EBITDA 率 × 債務上限（加權目標價，即時重算）`,
            tip: `每一格都以目前所有其他輸入，重新跑一次完整的資金瀑布與評價。列＝穩態 EBITDA 率、欄＝債務／backlog 上限；預設為 ${hA(DEFAULTS.ebSteady * 100, 0)} × ${Y(DEFAULTS.debtBacklog, 1)}x。`,
            w: 460
          }), (0, $.jsx)(`div`, {
            className: `mt-2`,
            style: xstyQ.wrap,
            children: (0, $.jsxs)(`table`, {
              style: {
                ...xstyQ.table,
                minWidth: 520
              },
              children: [(0, $.jsx)(`thead`, {
                children: (0, $.jsx)(`tr`, {
                  children: [`穩態 EBITDA 率 ＼ 債務上限`, `0.4x`, `1.0x`].map((e, t) => (0, $.jsx)(`th`, {
                    style: t === 0 ? xstyQ.thL : xstyQ.th,
                    children: e
                  }, e))
                })
              }), (0, $.jsx)(`tbody`, {
                children: [.35, .47, .59].map((q, z) => (0, $.jsxs)(`tr`, {
                  children: [(0, $.jsx)(`td`, {
                    style: {
                      ...xstyQ.tdL,
                      background: z % 2 ? `#fbfaf7` : `#fff`
                    },
                    children: hA(q * 100, 0)
                  }), [.4, 1].map(k => {
                    let st = {
                        ...t,
                        ebSteady: q,
                        debtBacklog: k
                      },
                      dd = runFunding(st),
                      pp = runValuation(dd, st, n),
                      on = Math.abs(q - t.ebSteady) < 1e-6 && Math.abs(k - t.debtBacklog) < 1e-6;
                    return (0, $.jsx)(`td`, {
                      style: {
                        ...xstyQ.td,
                        background: on ? `#fff2a8` : z % 2 ? `#fbfaf7` : `#fff`,
                        fontWeight: on ? 700 : 400
                      },
                      children: `$${Y(pp.call.blended, 1)}（股權 ${Y(dd.totals.equity, 0)}）`
                    }, k)
                  })]
                }, q))
              })]
            })
          }), (0, $.jsx)(hdrQ, {
            title: `錨定年度 × 倍數：EV/EBITDA 腿的方法敏感度（即時重算）`,
            tip: `EV/EBITDA 腿（分部加總）＝（錨定年 AI 雲端 EBITDA × 倍數＋非 AI 事業 EBITDA × ${VAL_DEFAULTS.legacyEvEbitda == null ? `各分部同業倍數（錨定年度加權）` : `${multTxt(VAL_DEFAULTS.legacyEvEbitda)}x`} − 錨定年末淨負債）÷ 錨定年末股數，${CALQ.evDiscText}（目標價時點）。矩陣只變動 AI 雲端 倍數；非 AI 事業倍數固定為同業 NTM 中位數。AI 雲端 6x 為可觀察 neocloud 穩態倍數上緣，套在利潤率仍在爬坡的 ${PERIODS[1]} 上並不一致。黃底為目前設定，綠底為不低於現價。`,
            w: 480
          }), (0, $.jsx)(EvGridQ, {
            st: t,
            o: n
          }), (0, $.jsx)(hdrQ, {
            title: `敏感度：市場以多少價格、吸收多少新股（加權目標價，即時重算）`,
            tip: `列＝股權發行折價（相對發行參考價），欄＝每年股權吸收上限（占現市值）；超出上限的部分改以高息債補足。左表為目前債務上限，右表固定為 0.4x（公司簡報揭露的當前比率）。在目前 1.0x 下債務承擔大部分缺口，股權需求小且集中在 FY29–30，所以對目標價影響有限；債務上限收緊到 0.4x 時，這兩個變數才成為主導。`,
            w: 480
          }), (0, $.jsx)(`div`, {
            className: `mt-2 grid gap-3 lg:grid-cols-2`,
            children: [t.debtBacklog, .4].map((K, KI) => (0, $.jsx)(`div`, {
              style: xstyQ.wrap,
              children: (0, $.jsxs)(`table`, {
                style: {
                  ...xstyQ.table,
                  minWidth: 460
                },
                children: [(0, $.jsx)(`thead`, {
                  children: (0, $.jsx)(`tr`, {
                    children: [`債務上限 ${Y(K,1)}x｜折價 ＼ 每年上限`, `10%`, `20%`, `50%`, `無上限`].map((e, t) => (0, $.jsx)(`th`, {
                      style: t === 0 ? xstyQ.thL : xstyQ.th,
                      children: e
                    }, e))
                  })
                }), (0, $.jsx)(`tbody`, {
                  children: [0, .1, .2, .3].map((q, z) => (0, $.jsxs)(`tr`, {
                    children: [(0, $.jsx)(`td`, {
                      style: {
                        ...xstyQ.tdL,
                        background: z % 2 ? `#fbfaf7` : `#fff`
                      },
                      children: hA(q * 100, 0)
                    }), [.1, .2, .5, 99].map(cp => {
                      let st = {
                          ...t,
                          debtBacklog: K,
                          eqDisc: q,
                          eqCapPct: cp
                        },
                        dd = runFunding(st),
                        pp = runValuation(dd, st, n),
                        on = KI === 0 && Math.abs(q - t.eqDisc) < 1e-6 && Math.abs(cp - t.eqCapPct) < 1e-6;
                      return (0, $.jsx)(`td`, {
                        style: {
                          ...xstyQ.td,
                          background: on ? `#fff2a8` : z % 2 ? `#fbfaf7` : `#fff`,
                          fontWeight: on ? 700 : 400
                        },
                        title: `股權 ${Y(dd.totals.equity,1)}bn、高息債 ${Y(dd.totals.junk,1)}bn`,
                        children: [`$${Y(pp.call.blended, 1)}`, (0, $.jsx)(`div`, {
                          style: {
                            fontSize: 10.5,
                            color: `#8a8f98`,
                            fontWeight: 400,
                            whiteSpace: `nowrap`
                          },
                          children: `股 ${Y(dd.totals.equity,1)}｜高 ${Y(dd.totals.junk,1)}`
                        })]
                      }, cp)
                    })]
                  }, q))
                })]
              })
            }, KI))
          }), (() => {
            // v3.4：左表若全部相同，說明原因（DCF 腿截斷／EV/EBITDA 腿只看錨定年）
            let G = [0, .1, .2, .3].flatMap(q => [.1, .2, .5, 99].map(cp => {
                let st = {
                    ...t,
                    eqDisc: q,
                    eqCapPct: cp
                  },
                  dd = runFunding(st),
                  pp = runValuation(dd, st, n);
                return {
                  tgt: pp.call.blended,
                  raw: pp.d.perShareRaw,
                  ev: pp.peAdj,
                  yrs: dd.years.filter(e => e.equity + e.junk > .05).map(e => e.year)
                }
              })),
              lo = Math.min(...G.map(e => e.tgt)),
              hi = Math.max(...G.map(e => e.tgt)),
              rl = Math.min(...G.map(e => e.raw)),
              rh = Math.max(...G.map(e => e.raw)),
              yrs = [...new Set(G.flatMap(e => e.yrs))].sort();
            return (0, $.jsx)(`p`, {
              className: `mt-2 text-xs leading-relaxed text-muted`,
              children: hi - lo < .05 ? `左表 16 格皆為 $${Y(lo,1)}，計算正確但此表在目前設定下不具鑑別力：(1) DCF 腿股權價值為負（未截斷每股 ${mA(rl,1)} 至 ${mA(rh,1)}，確實隨折價與上限變動），被 ${n.dcfMode === `option` ? `選擇權模式壓縮` : `0 截斷`}；(2) EV/EBITDA 腿以 ${PERIOD_LABELS[n.evYear ?? 1]} 末淨負債與股數計算，而股權與高息債只出現在 ${yrs.join('、') || '—'}${yrs.every(y => y > PERIOD_LABELS[n.evYear ?? 1]) ? `，不影響錨定年` : ``}。每格下方小字為五期股權｜高息債（$bn），可看出融資組合確有變化；右表（0.4x）股權需求提前，兩腿都會反映。` : `左表範圍 $${Y(lo,1)}–$${Y(hi,1)}；每格下方小字為五期股權｜高息債（$bn）。`
            })
          })(), (0, $.jsx)(`table`, {
            className: `mt-4 w-full text-sm`,
            children: (0, $.jsx)(`tbody`, {
              children: [
                [`DCF（已含 CapEx）`, c.perShareT, `不重複扣缺口；以 WACC 推到 ${CALQ.targetDate}`],
                [`EV/EBITDA（融資後）`, f, `${PERIOD_LABELS[r.evK]} 末淨負債 ${Y(r.ndA,1)}bn ／ ${PERIOD_LABELS[r.evK]} 末股數 ${Y(r.shA,3)}bn${r.evK >= CALQ.evDiscFrom ? `；以 WACC 折回 ${CALQ.targetText}` : ``}`],
              ].map(([e, t, n]) => (0, $.jsxs)(`tr`, {
                className: `border-t border-border`,
                children: [(0, $.jsx)(`td`, {
                  className: `py-2 font-medium`,
                  children: e
                }), (0, $.jsxs)(`td`, {
                  className: `py-2 text-right font-mono tabular-nums`,
                  children: [`$`, Y(Number(t), 0)]
                }), (0, $.jsx)(`td`, {
                  className: `py-2 text-right text-xs text-muted`,
                  children: n
                })]
              }, String(e)))
            })
          }), (0, $.jsxs)(`div`, {
            className: `mt-4 grid gap-3 md:grid-cols-3`,
            children: [(0, $.jsx)(MM, {
              title: `點位（評等依此）`,
              body: `$${Y(TR.pt,1)} · ${m.call}`,
              hint: `目前輸入的加權目標價。賣出門檻 $${Y(TR.th,1)}（現價 −${pctQ(SELL_TH)}）；上檔 ≥${pctQ(RATE_TH.buyUpsideMin)} 且終值占比 <${pctQ(RATE_TH.buyTvShareMax)}、股權需求不超過現市值 ${multTxt(RATE_TH.equityRaiseMaxMult)} 倍才給買進。`
            }), (0, $.jsx)(MM, {
              title: `情境區間（保守與積極情境）`,
              body: `$${Y(TR.A[0],1)}–$${Y(TR.A[1],1)}`,
              hint: `只換擴張力道、保留其餘手動調整：${TR.scen.map(s => `${s.name} $${Y(s.tgt,1)}（${s.call}）`).join(`、`)}。`
            }), (0, $.jsx)(MM, {
              title: TR.bLabel,
              body: `$${Y(TR.B[0],1)}–$${Y(TR.B[1],1)}`,
              hint: `目前輸入，EV/EBITDA 倍數換成 ${multTxt(TR.mults[0])}x 與 ${multTxt(TR.mults[1])}x；${TR.pos}。`
            })]
          }), m.notes.length ? (0, $.jsx)(`ul`, {
            className: `mt-4 space-y-1 text-xs leading-relaxed text-muted`,
            children: m.notes.map(e => (0, $.jsxs)(`li`, {
              children: [`— `, e]
            }, e))
          }) : null, (0, $.jsx)(`p`, {
            className: `mt-4 text-xs leading-relaxed text-muted`,
            children: `缺口處理：採期前融資瀑布——每期在需要前先融足，使期末現金不低於最低現金；依序動用未動用額度、資產層新債（總債務 ≤ 債務／backlog 上限）、股權（按現價折價發行）。新債利息進損益，新股進股數；DCF 以融資後股數計每股，並加回新股募得現金的現值；EV/EBITDA 用錨定年末（預設 ${PERIODS[VAL_DEFAULTS.evYear]}）淨負債與股數。舉債部分反映在錨定年末淨負債（與舊版扣缺口本金等價），股權部分反映在股數。 結論規則：股權募資 > 現市值 ${multTxt(RATE_TH.equityRaiseMaxMult)} 倍 → 禁止買進；加權目標價低於現價 ${pctQ(SELL_TH)} 以上，或股權需求過大且目標價低於現價 ${pctQ(RATE_TH.sellUpsideMaxIfEquityOver)} 以上 → 賣出；上檔 ≥${pctQ(RATE_TH.buyUpsideMin)} 且股權需求未超標才買進。這是研究框架，不是投資建議。市場共識目標價與差異見「市場共識」分頁。`
          })]
        })]
      })]
    })]
  })
}
