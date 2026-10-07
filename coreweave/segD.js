function zM() {
  let [e, t] = (0, v.useState)(PM), [n, r] = (0, v.useState)(0), [i, a] = (0, v.useState)(`summary`), [o, s] = (0, v.useState)(() => structuredClone(VAL_DEFAULTS)), [c, l] = (0, v.useState)(0), u = (0, v.useRef)(null), d = (0, v.useMemo)(() => runFunding(e), [e]), f = (0, v.useMemo)(() => runValuation(d, e, o), [d, e, o]), TR = (0, v.useMemo)(() => targetRange(d, e, o, f), [d, e, o, f]), p = (0, v.useMemo)(() => EM(f, d, e), [f, d, e]), m = (0, v.useMemo)(() => sensitivities(e, o), [e, o]), h = lA(e.cds), g = d.totals.end >= 0, _ = d.totals.operatingGap + e.cash >= 0, y = (0, v.useMemo)(() => aA(e, d), [e, d]), b = (0, v.useMemo)(() => rpoBridge(d), [d]), x = (0, v.useMemo)(() => sA(e, d), [e, d]), S = (0, v.useMemo)(() => cA(e), [e]), C = e => s(t => ({
    ...t,
    ...e
  })), w = e => t(t => ({
    ...t,
    ...e,
    scenario: e.scenario ?? `custom`
  })), T = e => t(t => ({
    ...t,
    scenario: e,
    a: {
      ...structuredClone(t.a),
      newLease: [...SCENARIOS[e].a.newLease]
    },
    mw31: SCENARIOS[e].mw31,
    m: {
      ...t.m,
      accepted: [...SCENARIOS[e].acc],
      billable: [...SCENARIOS[e].bil]
    }
  })), E = (e, n, r) => t(t => {
    let i = structuredClone(t.a);
    return e === `margin` ? i.margin = r : i[e][n] = r, {
      ...t,
      a: i,
      scenario: `custom`
    }
  }), D = (e, n, r) => t(t => {
    let i = structuredClone(t.m);
    return i[e][n] = r, {
      ...t,
      m: i,
      scenario: `custom`
    }
  }), O = (e, n, r) => t(t => {
    let i = structuredClone(t.sites),
      a = i[e];
    return !a || a.residual ? t : (a[n] = r, {
      ...t,
      sites: i,
      scenario: `custom`
    })
  }), k = (e, t) => {
    let n = new Blob([`﻿` + t.map(e => e.join(`,`)).join(`
`)], {
        type: `text/csv;charset=utf-8`
      }),
      r = document.createElement(`a`);
    r.href = URL.createObjectURL(n), r.download = e, r.click(), URL.revokeObjectURL(r.href)
  }, Rw = (e, t, n = 1) => (0, $.jsxs)(`tr`, {
    className: `border-b border-border/70`,
    children: [(0, $.jsx)(`td`, {
      className: `py-1.5 pr-3 font-sans font-medium`,
      children: e
    }), t.map((e, t) => (0, $.jsx)(`td`, {
      className: `py-1.5 text-right font-mono tabular-nums`,
      children: typeof e == `number` ? Y(e, n) : e
    }, t))]
  });
  return Math.max(...d.years.flatMap(e => [e.sources, e.uses]), 1), (0, $.jsxs)(`div`, {
    className: `min-h-screen overflow-x-hidden bg-bg text-fg`,
    children: [(0, $.jsx)(`header`, {
      className: `sumQ-np border-b border-border bg-ink text-accent-fg`,
      children: (0, $.jsxs)(`div`, {
        className: `mx-auto flex max-w-7xl flex-col gap-5 px-4 py-8 md:px-6`,
        children: [(0, $.jsxs)(`div`, {
          className: `flex flex-wrap items-center gap-2 text-xs tracking-wide text-accent-soft`,
          children: [(0, $.jsx)(`span`, {
            children: `${COMPANY_DATA.meta.ticker} ${PERIODS[0]}–${PERIODS[PERIODS.length - 1]}` // v4.5：期間由日曆推算
          }), (0, $.jsx)(`span`, {
            className: `text-subtle`,
            children: `/`
          }), (0, $.jsx)(`span`, {
            children: `產能約束 · 資金優先`
          }), (0, $.jsx)(`span`, {
            className: `text-subtle`,
            children: `/`
          }), (0, $.jsxs)(`span`, {
            children: [`v`, `1.0`, ` · `, UPDATE_DATE]
          }), (0, $.jsx)(`span`, {
            className: `text-subtle`,
            children: `/`
          }), (0, $.jsx)(`span`, {
            children: COMPANY_DATA.texts.sourceLine // v4.5：資料來源一行（company.json → texts，隨每季資料更新）
          })]
        }), (0, $.jsx)(`h1`, {
          className: `max-w-3xl font-display text-3xl font-semibold leading-tight tracking-tight md:text-4xl`,
          children: `Backlog 不是現金`
        }), (0, $.jsx)(`p`, {
          className: `max-w-3xl font-display text-lg font-medium leading-snug text-accent-fg md:text-xl`,
          children: `每 MW 賺的錢付不起 GPU 的資本成本——蓋得愈多，愈要靠外部資金`
        }), (0, $.jsxs)(`p`, {
          className: `max-w-3xl text-sm leading-relaxed text-pretty text-accent-soft`,
          children: [e.scenario === `custom` ? `自訂情境` : SCENARIOS[e.scenario]?.label, `：${PERIODS[0]}–${PERIODS[PERIODS.length - 1].slice(2)} 融資前缺口 `, (0, $.jsxs)(`span`, {
            className: `font-medium text-accent-fg`,
            children: [mA(Math.max(0, -d.totals.preFinEnd)), `bn`]
          }), `，加上新融資本身的利息 `, mA(d.years.reduce((e, t) => e + t.newDebtInt, 0)), `bn，由新債 `, mA(d.totals.newDebt), `bn、股權 `, mA(d.totals.equity), `bn（新股 `, Y(d.totals.newShares, 2), `bn 股）`, d.totals.junk > .05 ? `、高息債 ${mA(d.totals.junk)}bn` : ``, ` 補足。每 MW 年 EBITDA 約 `, (0, $.jsxs)(`span`, {
            className: `font-medium text-accent-fg`,
            children: [`$`, Y(d.m.revMW[4] * (e.revScale ?? 1) * 1e3 * d.m.util[4] / 100 * d.years[4].ebM, 1), `m`]
          }), `，回收一個 MW 的 GPU（$`, Y(e.a.costMW[4] * (e.capexScale ?? 1), 0), `m、`, e.gpuLife, ` 年、WACC `, hA(o.wacc * 100, 0), `）每年需要 `, (0, $.jsxs)(`span`, {
            className: `font-medium text-accent-fg`,
            children: [`$`, Y(e.a.costMW[4] * (e.capexScale ?? 1) * o.wacc / (1 - Math.pow(1 + o.wacc, -e.gpuLife)), 1), `m`]
          }), `。目標價 `, (0, $.jsxs)(`span`, {
            className: `font-medium text-accent-fg`,
            children: [`$`, Y(f.call.blended, 1)]
          }), `（情境區間 $${Y(TR.A[0], 1)}–$${Y(TR.A[1], 1)}），較現價 `, f.call.upside >= 0 ? `+` : ``, hA(f.call.upside * 100, 0), `，結論「`, f.call.call, `」。`, (0, $.jsx)(tipQ, {
            t: `三個判斷依據：(1) Backlog 不是現金——RPO 103.7bn 須先蓋出產能才能認列，五期認列 84%，產能不足時收不到。(2) 單位經濟——每 MW 年 EBITDA（每 MW 年收入×利用率×FY30 EBITDA 率）低於 GPU 的年化資本回收（每 MW 成本×資本回收係數），擴張本身不創造價值；要翻轉須每 MW 年收入約 $14m 以上或每 MW 成本降至約 $27m 以下。(3) 融資——缺口依序由額度、新債（總債務 ≤ 1.0× backlog，依 2026 年實際融資組合校準）、股權（每年上限＝現市值 20%）、高息債補足；市場願意以多少價格吸收多少新股，是估值最敏感的假設之一（見目標價頁敏感度表）。FY26 欄為 1H 實際＋2H 模型。`,
            w: 500
          })]
        }), (0, $.jsxs)(`div`, {
          className: `flex flex-wrap gap-2`,
          children: [(0, $.jsxs)(qj, {
            variant: `solid`,
            onClick: () => k(`crwv_annual_v10.csv`, [Object.keys(d.years[0]), ...d.years.map(e => Object.values(e).map(e => typeof e == `number` ? e : String(e)))]),
            children: [(0, $.jsx)(rM, {
              className: `size-4`
            }), `年度 CSV`]
          }), (0, $.jsx)(qj, {
            variant: `outline`,
            className: `border-white/20 bg-transparent text-accent-fg hover:bg-white/10`,
            onClick: () => k(`crwv_sites_v10.csv`, [Object.keys(d.sites[0]), ...d.sites.map(e => Object.values(e).map(e => String(e ?? ``)))]),
            children: `站點 CSV`
          }), (0, $.jsx)(qj, {
            variant: `outline`,
            className: `border-white/20 bg-transparent text-accent-fg hover:bg-white/10`,
            onClick: () => {
              let t = new Blob([JSON.stringify(e, null, 2)], {
                  type: `application/json`
                }),
                n = document.createElement(`a`);
              n.href = URL.createObjectURL(t), n.download = `crwv_model_v10.json`, n.click()
            },
            children: `儲存 JSON`
          }), (0, $.jsxs)(qj, {
            variant: `outline`,
            className: `border-white/20 bg-transparent text-accent-fg hover:bg-white/10`,
            onClick: () => u.current?.click(),
            children: [(0, $.jsx)(aM, {
              className: `size-4`
            }), `載入`]
          }), (0, $.jsx)(`input`, {
            ref: u,
            hidden: !0,
            type: `file`,
            accept: `application/json`,
            onChange: e => {
              let n = e.target.files?.[0];
              if (!n) return;
              let r = new FileReader;
              r.onload = () => {
                try {
                  t({
                    ...PM(),
                    ...JSON.parse(String(r.result))
                  })
                } catch {}
              }, r.readAsText(n)
            }
          }), (0, $.jsxs)(qj, {
            variant: `ghost`,
            className: `text-accent-fg hover:bg-white/10`,
            onClick: () => {
              t(PM()), s(structuredClone(VAL_DEFAULTS))
            },
            children: [(0, $.jsx)(iM, {
              className: `size-4`
            }), `重設`]
          })]
        })]
      })
    }), (0, $.jsxs)(`div`, {
      className: `sumQ-main mx-auto max-w-7xl space-y-4 px-4 py-5 md:px-6 md:py-6`,
      children: [(0, $.jsx)(`div`, {
        className: `sumQ-np flex w-full max-w-full overflow-x-auto`,
        children: (0, $.jsx)(`div`, {
          className: `flex gap-1 rounded-lg bg-surface p-1`,
          children: [
            [`summary`, `總結`],
            [`funding`, `資金模型`],
            [`value`, `損益與評價`]
          ].map(([e, t]) => (0, $.jsx)(`button`, {
            onClick: () => a(e),
            className: Vj(`h-11 min-w-28 rounded-md px-4 text-sm font-medium`, i === e ? `bg-ink text-accent-fg` : `text-muted hover:text-fg`),
            children: t
          }, e))
        })
      }), (0, $.jsx)(`div`, {
        className: Vj(i !== `summary` && `hidden`),
        children: (0, $.jsx)(SumQ, {
          d: d,
          f: f,
          e: e,
          o: o,
          m: m,
          tr: TR,
          active: i === `summary`
        })
      }), (0, $.jsx)(`div`, {
        className: Vj(i !== `value` && `hidden`),
        children: (0, $.jsx)(AM, {
          result: d,
          state: e,
          v: o,
          pack: f,
          tr: TR,
          onPatch: C,
          tab: c,
          onTab: l
        })
      }), (0, $.jsxs)(`div`, {
        className: Vj(i !== `funding` && `hidden`),
        children: [(0, $.jsx)(`div`, {
          className: `grid grid-cols-2 gap-3 md:grid-cols-3 lg:grid-cols-6`,
          children: [
            [`排程 RPO（五期）`, mA(b.scheduled) + `bn`],
            [`可實現（扣產能／信用）`, mA(b.collected) + `bn`],
            [`營運缺口（不含股權）`, mA(d.totals.operatingGap) + `bn`],
            [`期初＋可轉債＋瀑布融資`, mA(e.cash + d.totals.atm + d.totals.facility) + `bn`],
            [`需股權募資（瀑布）`, mA(d.totals.equity) + `bn`],
            [`連動評價／結論`, `$${Y(f.call.blended,1)}（$${Y(TR.A[0],1)}–$${Y(TR.A[1],1)}）· ${f.call.call}`]
          ].map(([e, t], n) => (0, $.jsxs)(Jj, {
            className: `p-3 md:p-4`,
            children: [(0, $.jsx)(`div`, {
              className: `text-xs text-muted`,
              children: e
            }), (0, $.jsx)(`div`, {
              className: Vj(`mt-1 font-display text-xl font-semibold tabular-nums`, n === 2 && (d.totals.operatingGap >= 0 ? `text-ok` : `text-bad`), n === 4 && (g ? `text-ok` : `text-bad`), n === 5 && (f.call.call === `買進` ? `text-ok` : f.call.call === `賣出` ? `text-bad` : `text-watch`)),
              children: t
            })]
          }, e))
        }), (0, $.jsxs)(`div`, {
          className: `grid gap-4 lg:grid-cols-[minmax(0,20rem)_minmax(0,1fr)]`,
          children: [(0, $.jsx)(`aside`, {
            className: `min-w-0 space-y-4`,
            children: (0, $.jsxs)(Jj, {
              children: [(0, $.jsx)(hdrQ, {
                title: `管理層情境與全域假設`,
                tip: `三個情境只改變「管理層要蓋多少」：2030 年 Accepted MW 分別為 4.2／5.6／8 GW。保守＝只交付已簽約電力 4.2 GW（8/11 法說）、不再新簽；基準＝2027 年後每年約 +800 MW；積極＝公司目標 ≥8 GW。Billable 依相同爬坡比例推算；毛 CapEx 由 MW 公式自動增減，表外租金依新增 MW 等比例縮放，FY31 預建分別為 0／500／1,000 MW。其餘假設三情境相同，差異完全歸因於擴張力道。切換情境只會改寫 Accepted／Billable MW、FY31 預建與表外租金；你手動調整的其他數字（如每 MW 建置成本、每 MW 年收入、客戶預付、JV 後續增資）會保留，按頁首「重設」才回到預設值。FY26–FY27 三情境相同（年底 1,850／3,100 MW），FY28 起分歧。`,
                w: 460
              }), (0, $.jsx)(`div`, {
                className: `mt-3 grid grid-cols-3 gap-1 rounded-lg bg-surface p-1`,
                children: [`low`, `base`, `high`].map(t => (0, $.jsx)(`button`, {
                  onClick: () => T(t),
                  className: Vj(`h-10 rounded-md text-sm font-medium`, e.scenario === t ? `bg-ink text-accent-fg` : `text-muted hover:bg-card`),
                  children: SCENARIOS[t].label
                }, t))
              }), (0, $.jsxs)(`p`, {
                className: `mt-2 text-xs text-muted`,
                children: [`目前：`, e.scenario === `custom` ? `自訂` : SCENARIOS[e.scenario]?.label, ` · 2030 年 `, Y(d.m.accepted[4], 0), ` MW · 模型期 CapEx `, Y(d.years.reduce((e, t) => e + t.gross, 0), 0), `bn`]
              }), (0, $.jsxs)(`div`, {
                className: `mt-4 space-y-2`,
                children: [(0, $.jsx)(accQ, {
                  title: `B｜產能與收入`,
                  sum: `RPO ${Y(e.rpoOpen,1)}＋${Y(e.rpoPendingAdd,0)} · 認列 ${e.rp}%`,
                  children: [(0, $.jsx)(LM, {
                    label: `每 MW 年收入倍數（整體）`,
                    hint: `${hA((e.revScale ?? 1)*100,0)} · 反向 DCF 與壓力測試用；預設 100%`,
                    children: (0, $.jsx)(IM, {
                      value: (e.revScale ?? 1) * 100,
                      onChange: e => w({
                        revScale: Math.max(10, e) / 100
                      }),
                      step: 5
                    })
                  }),
(0, $.jsx)(LM, {
                  label: `6/30 RPO（US$bn）`,
                  hint: `103.7 10-Q（backlog 104.2）`,
                  children: (0, $.jsx)(IM, {
                    value: e.rpoOpen,
                    onChange: e => w({
                      rpoOpen: e
                    })
                  })
                }),
(0, $.jsx)(LM, {
                  label: `Q3 初新增承諾（US$bn）`,
                  hint: `法說「>25」，未入 RPO；口徑為 net new commitments`,
                  children: (0, $.jsx)(IM, {
                    value: e.rpoPendingAdd,
                    onChange: e => w({
                      rpoPendingAdd: e
                    })
                  })
                }),
(0, $.jsx)(LM, {
                  label: `五期認列比例（至 2030 末）`,
                  hint: hA(e.rp, 0),
                  children: (0, $.jsx)(`input`, {
                    type: `range`,
                    min: 50,
                    max: 100,
                    value: e.rp,
                    onChange: e => w({
                      rp: Number(e.target.value)
                    }),
                    className: `w-full accent-accent`
                  })
                }),
(0, $.jsx)(LM, {
                  label: `6/30 Billable MW`,
                  hint: `Assumed · 主動 1,500 減爬坡中`,
                  children: (0, $.jsx)(IM, {
                    value: e.billableOpen,
                    onChange: e => w({
                      billableOpen: e
                    }),
                    step: 50
                  })
                }),
(0, $.jsx)(RM, {
                  checked: e.useAvgMw,
                  onChange: e => w({
                    useAvgMw: e
                  }),
                  label: `收入改用平均在役 MW`,
                  hint: `關掉則用期末存量全期化。打開用（6/30 ${e.billableOpen} + 當期期末）/2。CoreWeave 交付後載（6 月 300 MW），預設打開。`
                }),
(0, $.jsx)(LM, {
                  label: `${PERIOD_FY[0] - 1} 年底主動電力 MW`,
                  hint: `Q4 法說 >850 [Verified]`,
                  children: (0, $.jsx)(IM, {
                    value: e.mwYearEnd[PERIOD_FY[0] - 1],
                    onChange: v => w({
                      mwYearEnd: { ...e.mwYearEnd, [PERIOD_FY[0] - 1]: v }
                    }),
                    step: 10
                  })
                }),
(0, $.jsx)(RM, {
                  checked: e.linkSites,
                  onChange: t => w({
                    linkSites: t,
                    scenario: e.scenario
                  }),
                  label: `站點 MW 連動年度容量`,
                  hint: `2H26 Accepted/Billable 不低於具名站點加總；不足部分進殘差列。`
                })]
                }),
                (0, $.jsx)(accQ, {
                  title: `C｜利潤率（EBITDA 單一來源）`,
                  sum: `EBITDA ${hA(ebPathQ(e, d)[0]*100,0)}→${hA(ebPathQ(e, d)[1]*100,0)}`,
                  children: [(0, $.jsx)(LM, {
                  label: `起始 EBITDA 率（FY26）`,
                  hint: `${hA(e.ebStart*100,1)} · Q2 實際 Adj. EBITDA 率 58.6%`,
                  children: (0, $.jsx)(IM, {
                    value: e.ebStart * 100,
                    onChange: e => w({
                      ebStart: e / 100
                    }),
                    step: .5
                  })
                }),
(0, $.jsxs)(LM, {
                  label: `穩態 EBITDA 率（FY30）`,
                  hint: `${hA(ebPathQ(e, d)[1]*100,1)} · ${PMWQ.cost === `bottomUp` ? `由下而上（改數值＝FY30 目標，差額線性分攤）` : `線性爬升`}；敏感度 59%／70%`,
                  children: [(0, $.jsx)(IM, {
                    value: ebPathQ(e, d)[1] * 100,
                    onChange: e => w({
                      ebSteady: e / 100
                    }),
                    step: .5
                  }), (0, $.jsxs)(`p`, {
                    className: `text-xs text-muted`,
                    children: [`損益與資金共用；資金端用 EBITDAR 率＝EBITDA 率＋租金÷營收`, (0, $.jsx)(tipQ, {
                      t: `由下而上：每 MW 年收入約 $10.9m，電費約 $0.70m、機房租金 $1.4–1.8m、維護頻寬約 $0.6m、現金 SG&A／R&D 約 $0.5m，合計約 29–33%，對應 EBITDA 率 67–71%。Q2 實際 59% 含爬坡期已起租未計費產能的成本。基準取 65%（中位數），59% 為維持現況、70% 為上緣。EBITDA 已扣營業租賃成本；資金模型把租金列在支出端，所以 EBITDAR 率須把租金加回，兩邊才不會重複扣除。`,
                      w: 440
                    })]
                  })]
                }),
(0, $.jsx)(RM, {
                  checked: e.overlay,
                  onChange: e => w({
                    overlay: e
                  }),
                  label: `另扣電力＋維護 overlay`,
                  hint: `貢獻率若已淨額化，請保持關閉。`
                })]
                }),
                (0, $.jsx)(accQ, {
                  title: `D｜支出（CapEx／折舊）`,
                  sum: `$${Y(e.a.costMW[0],0)}m/MW · λ ${hA(e.lambda*100,0)} · 壽命 ${e.gpuLife} 年`,
                  children: [(0, $.jsx)(LM, {
                    label: `每 MW 建置成本倍數（整體）`,
                    hint: `${hA((e.capexScale ?? 1)*100,0)} · 同時影響 CapEx、汰換與車隊折舊；預設 100%`,
                    children: (0, $.jsx)(IM, {
                      value: (e.capexScale ?? 1) * 100,
                      onChange: e => w({
                        capexScale: Math.max(10, e) / 100
                      }),
                      step: 5
                    })
                  }),
(0, $.jsx)(LM, {
                  label: `CapEx 提前支出比例 λ`,
                  hint: `${Y(e.lambda*100,0)}% · 次年上線 MW 的支出落在前一年的比例`,
                  children: (0, $.jsx)(IM, {
                    value: e.lambda * 100,
                    onChange: e => w({
                      lambda: e / 100
                    }),
                    step: 5
                  })
                }),
(0, $.jsx)(LM, {
                  label: `FY31 新增 MW（FY30 預建用）`,
                  hint: `0＝2030 達 8 GW 後停；1,700＝維持步調`,
                  children: (0, $.jsx)(IM, {
                    value: e.mw31,
                    onChange: e => w({
                      mw31: e
                    }),
                    step: 100
                  })
                }),
(0, $.jsx)(LM, {
                  label: `FY26 CapEx 下限（已承諾）`,
                  hint: `全年指引下緣 35；保守情境會觸發`,
                  children: (0, $.jsx)(IM, {
                    value: e.capexFloorFY0,
                    onChange: e => w({
                      capexFloorFY0: e
                    }),
                    step: 1
                  })
                }),
(0, $.jsx)(LM, {
                  label: `GPU 經濟壽命（年）`,
                  hint: `公司折舊年限 6 年；決定汰換時點與車隊折舊`,
                  children: (0, $.jsx)(IM, {
                    value: e.gpuLife,
                    onChange: e => w({
                      gpuLife: Math.max(3, Math.round(e))
                    }),
                    step: 1
                  })
                }),
(0, $.jsx)(LM, {
                  label: `${Object.keys(e.mwYearEnd)[0]} 年底主動電力 MW`,
                  hint: `[Assumed] 由 2023 CapEx 約 3.1bn ÷ $30m 推估`,
                  children: (0, $.jsx)(IM, {
                    value: Object.values(e.mwYearEnd)[0],
                    onChange: v => w({
                      mwYearEnd: { ...e.mwYearEnd, [Object.keys(e.mwYearEnd)[0]]: v }
                    }),
                    step: 10
                  })
                })]
                }),
                (0, $.jsx)(accQ, {
                  title: `E｜融資（瀑布與信用）`,
                  sum: `上限 ${Y(e.debtBacklog,1)}x · 股權 ${e.eqCapPct>=9?`無上限`:hA(e.eqCapPct*100,0)} · CDS ${Y(e.cds,0)}`,
                  children: [(0, $.jsx)(LM, {
                  label: `期初現金 6/30（US$bn）`,
                  hint: `5.524 + 有價證券 0.015；受限現金 1.38 不計`,
                  children: (0, $.jsx)(IM, {
                    value: e.cash,
                    onChange: e => w({
                      cash: e
                    })
                  })
                }),
(0, $.jsx)(RM, {
                  checked: e.includeAtm,
                  onChange: t => w({
                    includeAtm: t,
                    scenario: e.scenario
                  }),
                  label: `計入 9/17 可轉債淨現金 $3.2bn ＋ ATM 上限約 $3.0bn`,
                  hint: `可轉債 3.7bn 定價（9/22 交割），扣承銷費後淨額約 3.6bn，再扣 capped call 成本約 0.45bn（比照 Q2 4.0bn 可轉債的 0.492bn）→ 可用現金約 3.2bn；ATM 35m 股 × $85.43（9/21 收盤；v4.3 現價更新後此金額維持不變）≈ 3.0bn 為上限，實際未必用盡。合計 ${e.atm}bn 記於 2H26，不算進「營運缺口」。可轉債同時是負債（2.875%，2033 到期，不在五期還本表內）。`
                }),
(0, $.jsx)(LM, {
                  label: `股權／可轉債金額（US$bn）`,
                  hint: `3.2（可轉債淨現金）＋ 3.0（ATM 上限）`,
                  children: (0, $.jsx)(IM, {
                    value: e.atm,
                    onChange: e => w({
                      atm: e
                    })
                  })
                }),
(0, $.jsx)(RM, {
                  checked: e.useFacility,
                  onChange: t => w({
                    useFacility: t,
                    scenario: e.scenario
                  }),
                  label: `瀑布可動用未動用額度 $${e.facility}bn`,
                  hint: `10-Q 可用額度 10.014（RCF＋DDTL 未動用；DDTL 4.0 可動用至 2027-06-30）。瀑布第一順位，不受 債務／backlog 上限限制（已承諾額度）。關閉＝不動用額度，直接進入資產層新債與股權。`
                }),
(0, $.jsx)(RM, {
                  checked: e.includeDebt,
                  onChange: e => w({
                    includeDebt: e
                  }),
                  label: `債務排程攤還（10-Q 本金表）`,
                  hint: `DDTL 隨客戶付款攤還、OEM 融資短期到期，屬契約性支出，預設打開。關閉＝假設全額再融資。不含新借款。`
                }),
(0, $.jsx)(LM, {
                  label: `債務／backlog 上限`,
                  hint: `預設 1.0x：依 2026 年實際融資組合（股權約 23%）校準；0.4x 為公司 9/17 簡報揭露的當前比率`,
                  children: (0, $.jsx)(IM, {
                    value: e.debtBacklog,
                    onChange: e => w({
                      debtBacklog: e
                    }),
                    step: .05
                  })
                }),
(0, $.jsx)(LM, {
                  label: `新簽合約年期（年）`,
                  hint: `Q4 法說：加權平均約 5 年；決定 backlog 補入量`,
                  children: (0, $.jsx)(IM, {
                    value: e.ctrTerm,
                    onChange: e => w({
                      ctrTerm: e
                    }),
                    step: 1
                  })
                }),
(0, $.jsx)(LM, {
                  label: `最低現金（US$bn）`,
                  hint: `期前融資的現金底線 [Assumed]`,
                  children: (0, $.jsx)(IM, {
                    value: e.minCash,
                    onChange: e => w({
                      minCash: e
                    }),
                    step: .5
                  })
                }),
(0, $.jsx)(LM, {
                  label: `股權發行價（US$）`,
                  hint: `預設 ${mdQ(PRICE_DATE)} 收盤 ${Y(COMPANY_DATA.defaults.eqPx, 2)}`,
                  children: (0, $.jsx)(IM, {
                    value: e.eqPx,
                    onChange: e => w({
                      eqPx: e
                    }),
                    step: 1
                  })
                }),
(0, $.jsx)(LM, {
                  label: `股權發行折價 %`,
                  hint: `大額增資的折讓 [Assumed]`,
                  children: (0, $.jsx)(IM, {
                    value: e.eqDisc * 100,
                    onChange: e => w({
                      eqDisc: e / 100
                    }),
                    step: 1
                  })
                }),
(0, $.jsx)(LM, {
                  label: `每年股權吸收上限（占現市值 %）`,
                  hint: e.eqCapPct >= 9 ? `無上限` : `${hA(e.eqCapPct*100,0)} · 約 $${Y(e.eqCapPct*e.eqPx*.5865,1)}bn／年；2026 年實際約 13%`,
                  children: (0, $.jsx)(IM, {
                    value: e.eqCapPct * 100,
                    onChange: e => w({
                      eqCapPct: e / 100
                    }),
                    step: 5
                  })
                }),
(0, $.jsx)(LM, {
                  label: `高息債利率（股權上限溢出）`,
                  hint: `${hA(e.junkRate*100,1)} · 高於 2031 票據 9.75%；CDS 720bps [Assumed]`,
                  children: (0, $.jsx)(IM, {
                    value: e.junkRate * 100,
                    onChange: e => w({
                      junkRate: e / 100
                    }),
                    step: .5
                  })
                }),
(0, $.jsx)(RM, {
                  checked: e.cdsLink,
                  onChange: t => w({
                    cdsLink: t,
                    scenario: e.scenario
                  }),
                  label: `CDS 溢價傳入新債利率`,
                  hint: `新債利率 = 基準 9% + max(0, CDS−450)×40%。基準 9% 已對應 DDTL 5.0／高收益債水準，不要再開，避免雙計。不進入客戶違約。`
                }),
(0, $.jsxs)(`div`, {
                className: `mt-5 rounded-lg border border-border bg-surface p-3`,
                children: [(0, $.jsx)(`div`, {
                  className: `text-sm font-medium`,
                  children: `CoreWeave 5Y CDS`
                }), (0, $.jsxs)(`p`, {
                  className: `mt-1 text-xs leading-relaxed text-muted`,
                  children: [`中價 `, Y(e.cds, 0), ` bps（`, e.cdsDate, `）；Bid/Ask `, Y(e.cdsBid, 0), `／`, Y(e.cdsAsk, 0), `；近期區間 `, Y(e.cdsLo, 0), `–`, Y(e.cdsHi, 0), `（7/28 峰值 855、6 月低點 452）。這是 CoreWeave 自身信用，不是客戶違約率。`]
                }), (0, $.jsx)(LM, {
                  label: `CDS（bps）`,
                  hint: `${Y(e.cds,0)}`,
                  children: (0, $.jsx)(`input`, {
                    type: `range`,
                    min: 200,
                    max: 1200,
                    value: e.cds,
                    onChange: e => w({
                      cds: Number(e.target.value)
                    }),
                    className: `w-full accent-accent`
                  })
                }), (0, $.jsxs)(`div`, {
                  className: `mt-2 flex items-center justify-between`,
                  children: [(0, $.jsx)(Uj, {
                    tone: h.tone,
                    children: h.label
                  }), (0, $.jsxs)(`span`, {
                    className: `text-xs text-muted`,
                    children: [450, ` / `, 600, ` / `, 800]
                  })]
                })]
              })]
                })]
              })]
            })
          }), (0, $.jsxs)(`main`, {
            className: `min-w-0 space-y-4`,
            children: [(0, $.jsxs)(Jj, {
              children: [(0, $.jsxs)(`div`, {
                className: `flex flex-wrap items-center justify-between gap-2`,
                children: [(0, $.jsx)(`h2`, {
                  className: `font-display text-lg font-semibold`,
                  children: [`各期收支狀況`, (0, $.jsx)(tipQ, {
                    t: `左圖：各期營運來源、融資來源與用途。右圖：紅色面積為「融資前累積現金」＝若不做任何新融資，6/30 現金加上各期營運來源與可轉債／ATM、減去所有支出後的累積結果，負值即需要外部資金的規模；線條為瀑布累計補入的新債、股權與高息債。期前融資後的期末現金固定在最低現金 2.0bn，因此不再畫出。`,
                    w: 440
                  })]
                }), (0, $.jsxs)(`div`, {
                  className: `flex flex-wrap gap-2`,
                  children: [(0, $.jsxs)(Uj, {
                    tone: _ ? `ok` : `watch`,
                    children: [`營運`, _ ? `可覆蓋（含期初現金）` : `仍缺資金`]
                  }), (0, $.jsx)(Uj, {
                    tone: g ? `ok` : `bad`,
                    children: d.totals.equity > .01 ? `需股權 ${Y(d.totals.equity,0)}bn（新股 ${Y(d.totals.newShares,2)}bn 股）` : `缺口可全由債務支應`
                  })]
                })]
              }), (0, $.jsxs)(`div`, {
                className: `mt-4 grid gap-4 lg:grid-cols-2`,
                children: [(0, $.jsxs)(`div`, {
                  children: [(0, $.jsx)(`p`, {
                    className: `mb-2 text-xs text-muted`,
                    children: `來源分層：期初 RPO 現金／新簽約＋非算力現金／客戶預付／股權可轉債。股權可轉債不是營運來源。`
                  }), (0, $.jsx)(`div`, {
                    className: `h-56`,
                    children: (0, $.jsx)(Gs, {
                      width: `100%`,
                      height: `100%`,
                      children: (0, $.jsxs)(Lk, {
                        data: d.years.map(e => ({
                          name: e.year,
                          RPO現金: Number(e.rpoCash.toFixed(1)),
                          新簽約: Number((e.newCash + e.legacy).toFixed(1)),
                          客戶預付: Number(e.external.toFixed(1)),
                          股權可轉債: Number(e.atm.toFixed(1)),
                          用途: Number(e.uses.toFixed(1))
                        })),
                        children: [(0, $.jsx)(uD, {
                          stroke: `#ddd6c8`,
                          strokeDasharray: `3 3`
                        }), (0, $.jsx)(XD, {
                          dataKey: `name`,
                          tick: {
                            fontSize: 11
                          }
                        }), (0, $.jsx)(pO, {
                          tick: {
                            fontSize: 11
                          }
                        }), (0, $.jsx)(Es, {}), (0, $.jsx)(ow, {
                          dataKey: `RPO現金`,
                          stackId: `s`,
                          fill: `#0f5c61`
                        }), (0, $.jsx)(ow, {
                          dataKey: `新簽約`,
                          stackId: `s`,
                          fill: `#3d8b8f`
                        }), (0, $.jsx)(ow, {
                          dataKey: `客戶預付`,
                          stackId: `s`,
                          fill: `#8fb9bb`
                        }), (0, $.jsx)(ow, {
                          dataKey: `股權可轉債`,
                          stackId: `s`,
                          fill: `#c4a35a`
                        }), (0, $.jsx)(ow, {
                          dataKey: `用途`,
                          fill: `#5b5778`
                        })]
                      })
                    })
                  })]
                }), (0, $.jsx)(`div`, {
                  className: `h-56`,
                  children: (0, $.jsx)(Gs, {
                    width: `100%`,
                    height: `100%`,
                    children: (0, $.jsxs)(Rk, {
                      data: (() => {
                        let t = 0,
                          n = 0,
                          r = 0;
                        return d.years.map(e => (t += e.newDebt, n += e.equity, r += e.junk, {
                          name: e.year,
                          融資前累積現金: Number(e.preFinCum.toFixed(1)),
                          累計新債: Number(t.toFixed(1)),
                          累計股權: Number(n.toFixed(1)),
                          累計高息債: Number(r.toFixed(1))
                        }))
                      })(),
                      children: [(0, $.jsx)(uD, {
                        stroke: `#ddd6c8`,
                        strokeDasharray: `3 3`
                      }), (0, $.jsx)(XD, {
                        dataKey: `name`,
                        tick: {
                          fontSize: 11
                        }
                      }), (0, $.jsx)(pO, {
                        tick: {
                          fontSize: 11
                        }
                      }), (0, $.jsx)(Es, {}), (0, $.jsx)(ND, {
                        type: `monotone`,
                        dataKey: `融資前累積現金`,
                        stroke: `#9f1239`,
                        fill: `#f6d9df`
                      }), (0, $.jsx)(ND, {
                        type: `monotone`,
                        dataKey: `累計新債`,
                        stroke: `#0f5c61`,
                        fill: `transparent`
                      }), (0, $.jsx)(ND, {
                        type: `monotone`,
                        dataKey: `累計股權`,
                        stroke: `#c4a35a`,
                        fill: `transparent`
                      }), (0, $.jsx)(ND, {
                        type: `monotone`,
                        dataKey: `累計高息債`,
                        stroke: `#5b5778`,
                        fill: `transparent`
                      })]
                    })
                  })
                })]
              })]
            }), (0, $.jsx)(navQ, {
              groups: FMODS,
              cur: n,
              set: r
            }), (0, $.jsxs)(Jj, {
              className: `max-w-full overflow-x-auto`,
              children: [n === 0 && (0, $.jsxs)(`div`, {
                className: `space-y-3`,
                children: [(0, $.jsx)(hdrQ, {
                  title: `收支假設（FY26 欄＝1H 實際＋2H 模型）`,
                  tip: `支出五層：① 毛 CapEx（含 OEM 融資的非現金認列；客戶預付按比率抵減）→ ② 租金（在帳＝10-Q 到期表固定值；表外＝未起租 35.5＋單站上限 14.7＋355 MW 按造價計租的現金路徑）→ ③ 利息（存量債務依既有債務推算；瀑布新債另計利息）→ ④ JV／策略投資出資（10-Q 已承諾 1.7bn）→ ⑤ 排程還本（10-Q 本金表）。藍色可編輯格為輸入；灰底列為計算結果。本頁輸入為 2H26 半年金額，FY26 欄在「各期收支」頁才加回 1H 實際。`
                }), (0, $.jsx)(BM, {
                  rows: [
                    [`支出（可編輯輸入；FY26 欄為 2H 模型值）`, null],
                    [`　每 MW 建置成本`, e.a.costMW, (e, t) => E(`costMW`, e, t), 1, void 0, `GPU＋網路＋機房內裝（機房外殼由房東出資，走租金）。以 FY26 校準：YE25 850→YE26 1,850 MW＋預建 FY27，重現全年指引 35–39。FY26 指引把它釘在約 $32–37m（λ 25–45%）。`, `US$m/MW`],
                    [`　本期新增 MW`, d.years.map(e => e.mwNew), void 0, void 0, `calc`, `＝期末 Accepted − 期初。FY26 期初為 2025 年底 850 MW。`, `MW`],
                    [`　次期新增 MW`, d.years.map(e => e.mwNext), void 0, void 0, `calc`, `下一期新增量；FY30 欄取左欄「FY31 新增 MW」。`, `MW`],
                    [`　全年毛 CapEx（公式）`, d.years.map(e => e.capexFull), void 0, void 0, `calc`, `＝(本期新增×(1−λ)＋次期新增×λ)×每 MW 成本。FY26 應落在公司全年指引 35–39。`],
                    [`　成長型 CapEx（模型期）`, d.years.map(e => e.capexGrowth), void 0, void 0, `calc`, `FY26 欄＝MAX(全年公式, FY26 下限 ${e.capexFloorFY0}) − 1H 已認列 16.139。下限代表 2026 年已下單、無論情境都會發生的支出；保守情境的公式值低於下限，差額即「轉向太晚」的成本。`],
                    [`　GPU 汰換 CapEx`, d.years.map(e => e.refresh), void 0, void 0, `calc`, `＝(本年 − 經濟壽命) 那一年新增的 MW × 每 MW 成本。壽命 ${e.gpuLife} 年：FY29 汰換 2023 年批次（${Object.values(e.mwYearEnd)[0]} MW，推估）、FY30 汰換 2024 年批次（${Object.values(e.mwYearEnd)[1] - Object.values(e.mwYearEnd)[0]} MW；YE24 ${Object.values(e.mwYearEnd)[1]} MW 為 S-1 揭露）。三情境相同。`],
                    [`① 毛 CapEx（模型期）`, d.years.map(e => e.gross), void 0, void 0, `tot`, `＝成長型＋汰換。含 OEM 融資的非現金部分。`],
                    [`　對照：v1.4 手動值`, d.years.map(e => e.capexOld), void 0, void 0, void 0, `舊版未與 MW 連動，FY28–30 對每年 +1.5–1.7 GW 明顯不足。`],
                    [`　客戶預付率`, e.a.customerFund.map(e => e * 100), (e, t) => E(`customerFund`, e, t / 100), void 0, void 0, `H1 遞延收入淨流入 1.365 ÷ H1 CapEx 16.139 ＝ 8.5%。只降當期現金流出、形成遞延收入，不降專案總成本。`, `%`],
                    [`② 表外現金租金（未起租）`, e.a.newLease, (e, t) => E(`newLease`, e, t), void 0, void 0, `對應 10-Q 未起租租賃 35.5＋單站上限 14.7＋355 MW 按造價計租的現金支付路徑，全為假設。`],
                    [`② 在帳現金租金（10-Q 固定）`, [...LEASE_CASH_ON_BAL], void 0, void 0, void 0, `10-Q 到期表：2026 餘 1.04／27 2.34／28 2.31／29 2.37／30 2.29，2030 後尚有 19.02。`],
                    [`③ 存量債務利息（既有債務推算）`, d.years.map(e => e.intStock), void 0, void 0, `calc`, `＝平均本金（依 10-Q 到期表遞減）× 加權有效利率 8.4% × 期間長度 ＋ 9/17 可轉債利息 ＋ FY26 校準 0.35。明細見「既有債務」分頁。`],
                    [`　對照：v1.4 手動值`, d.years.map(e => e.intOld), void 0, void 0, void 0, `舊值未隨還本遞減，FY27 後每年高估 0.2–0.3。`],
                    [`③ 新債利息（瀑布，計算）`, d.years.map(e => e.newDebtInt), void 0, void 0, `calc`, `＝新債利率 × 期間長度 × (期初新債餘額 ＋ 本期舉借)。期前融資：本期舉借在期初到位，當期全額計息。`],
                    [`④ JV 已承諾餘額（10-Q）`, d.years.map(e => e.jvC), void 0, void 0, void 0, `10-Q Note 3：兩個 JV 承諾最高 1.7bn，預計 2026 年內履行；H1 已付 0.55 → 餘 1.15 全落在 FY26 下半年。`],
                    [`④ JV 後續增資＋策略投資`, e.a.div, (e, t) => E(`div`, e, t), void 0, void 0, `10-Q：入夥後須依協議追加出資，金額未揭露 [Assumed]。H1 策略投資 0.138。CRWV 不發股息。`],
                    [`⑤ 排程還本（10-Q 固定）`, e.includeDebt ? [...DEBT_AMORT] : DEBT_AMORT.map(() => 0), void 0, void 0, void 0, `DDTL 隨客戶付款攤還、OEM 融資短期到期。關閉開關＝假設全額再融資。`],
                    [`來源`, null],
                    [`非算力服務營收`, e.services, (n, r) => w({
                      services: e.services.map((e, t) => t === n ? r : e)
                    }), void 0, void 0, `軟體／儲存／CPU／managed inference。Q2 其他服務 ARR >0.4bn、推論 ARR 年底 ≥0.25bn。損益營收與資金現金共用這一列。`],
                    [`EBITDA 率（路徑）`, d.years.map(e => e.ebM * 100), void 0, void 0, `calc`, `由左欄起始與穩態值線性推得。`, `%`],
                    [`EBITDAR 率（計算）`, d.years.map(e => e.cashMargin * 100), void 0, void 0, `calc`, `＝EBITDA 率＋租金÷營收，即租金前的 EBITDA 率（業界稱 EBITDAR）。租金在支出端另列，故須加回。`, `%`],
                    [`非算力服務現金（計算）`, d.years.map(e => e.legacy), void 0, void 0, `calc`, `＝服務營收 × EBITDAR 率。`],
                    [`新簽約現金（計算）`, d.years.map(e => e.newCash), void 0, void 0, `calc`, `＝未被期初 RPO 占用的產能 × 新產能簽約率 × (1−信用損失率) × EBITDAR 率。FY28 起才出現。`]
                  ]
                })]
              }), n === 1 && (0, $.jsxs)(`div`, {
                className: `space-y-3`,
                children: [(0, $.jsx)(hdrQ, {
                  title: `各期收支（類現金流量）`,
                  tip: `FY26 欄＝1H26 實際（10-Q）＋2H26 模型，所以可直接對照公司全年指引：營收 12.4–13.2、CapEx 35–39。FY27 以後為純模型。負數以括號表示。現金橋：期初 ＋ 營運缺口 ＋ 股權／可轉債 ＋ 未動用額度 − 排程還本 ＝ 期末。`
                }), (0, $.jsx)(BM, {
                  rows: [
                    [`收入（產能約束）`, null],
                    [`排程 RPO`, d.years.map(e => e.scheduled), void 0, void 0, void 0, `＝(6/30 RPO 103.7 ＋ Q3 新增 25) × 本期桶權重 ÷ 權重合計 × 五期認列比例。只看合約，不看機房。`],
                    [`容量上限`, d.years.map(e => e.capacity), void 0, void 0, void 0, `＝平均在役 MW × 每 MW 年收入 × 利用率 × 期間係數。只看機房，不看合約。`],
                    [`平均在役 MW`, d.years.map(e => e.avgBillable), void 0, void 0, void 0, `＝(期初 Billable ＋ 期末 Billable) ÷ 2。可在左欄關閉，改用期末存量全期化。`, `MW`],
                    [`產能瓶頸`, d.years.map(e => e.bottleneck), void 0, void 0, `calc`, `＝MAX(0, 排程 − 容量)。合約有、機房沒有，本期收不到且不遞延到下期（保守處理）。`],
                    [`期初 RPO 轉換收入`, d.years.map(e => e.revenue), void 0, void 0, void 0, `＝排程 − 瓶頸。瓶頸>0 時恆等於容量上限。`],
                    [`新簽約收入（計算）`, d.years.map(e => e.newRev), void 0, void 0, `calc`, `＝MAX(0, 容量 − 排程) × 新產能簽約率。FY28 起模型收入主要來自這裡——尚未簽署的合約。`],
                    [`未售產能`, d.years.map(e => e.unsold), void 0, void 0, void 0, `蓋好但賣不掉。Bernstein：簽約電力 74% 位於 Tier 3/4 市場。`],
                    [`信用損失`, d.years.map(e => e.loss), void 0, void 0, void 0, `＝收入 × 違約率 × (1−回收率)。前三大客戶占 72% 營收。`],
                    [`損益用算力收入`, d.years.map(e => e.isRev), void 0, void 0, `tot`, `＝RPO 轉換 ＋ 新簽約。損益與評價頁用的是同一個數字。`],
                    [`來源（FY26 欄＝1H 實際現金流＋下半年模型）`, null],
                    [`Ⓐ0 1H26 實際營運現金流（CFO）`, d.years.map((e, t) => t === 0 ? e.fyCfo : 0), void 0, void 0, void 0, `10-Q 實際值 3.663，已含 1H 的利息、租金與客戶預付，因此下方②③Ⓓ的 FY26 欄只含下半年。`],
                    [`　EBITDA 率（損益、資金共用）`, d.years.map(e => e.ebM * 100), void 0, void 0, void 0, `起始 → 穩態線性爬升。`, `%`],
                    [`　EBITDAR 率（EBITDA 率＋租金÷營收）`, d.years.map(e => e.cashMargin * 100), void 0, void 0, void 0, `租金前的 EBITDA 率。租金在支出②另列，所以這裡加回，避免重複扣除。`, `%`],
                    [`Ⓐ RPO 現金（下半年起）`, d.years.map(e => e.rpoCash), void 0, void 0, void 0, `＝收現 × EBITDAR 率。`],
                    [`Ⓑ 新簽約現金`, d.years.map(e => e.newCash)],
                    [`Ⓒ 非算力服務現金`, d.years.map(e => e.legacy)],
                    [`Ⓓ 客戶預付（下半年起）`, d.years.map(e => e.external), void 0, void 0, void 0, `1H 的遞延收入淨流入 1.365 已含在 CFO 內，不重複計入。`],
                    [`營運來源合計`, d.years.map((e, t) => t === 0 ? e.fySourcesOp : e.sourcesOp), void 0, void 0, `tot`],
                    [`Ⓔ 股權／可轉債（融資）`, d.years.map((e, t) => t === 0 ? e.fyEquity : e.atm), void 0, void 0, void 0, `FY26＝1H 私募股權 2.982 ＋ 9/17 可轉債淨現金 3.2 ＋ ATM 上限 3.0。capped call 另列於用途⑥。`],
                    [`Ⓕ 1H26 實際借款（融資）`, d.years.map((e, t) => t === 0 ? e.fyBorrow : 0), void 0, void 0, void 0, `10-Q：H1 借款 16.747（還款另列於⑤）。1H 的缺口實際上就是靠舉債填補的。`],
                    [`Ⓖ 瀑布：新債（額度＋資產層）`, d.years.map(e => e.newDebt), void 0, void 0, void 0, `先用未動用額度，再用資產層新債；總債務不得超過 債務／backlog 上限。`],
                    [`Ⓗ 瀑布：股權募資`, d.years.map(e => e.equity), void 0, void 0, void 0, `債務用罄後的殘差，按現價折價發行；每年不超過股權吸收上限。`],
                    [`Ⓘ 瀑布：高息債（股權上限溢出）`, d.years.map(e => e.junk), void 0, void 0, void 0, `股權超過每年吸收上限的部分，以高息債補足；不受 backlog 上限約束。`],
                    [`總來源（含融資）`, d.years.map((e, t) => t === 0 ? e.fySrcTotal : e.sources), void 0, void 0, `tot`],
                    [`用途（現金口徑）`, null],
                    [`① CapEx（1H 現金購置／下半年起毛額）`, d.years.map((e, t) => t === 0 ? e.fyCapexUse : e.gross), void 0, void 0, void 0, `1H 採現金流量表的現金購置 14.117；下半年起採毛額，客戶預付在來源端抵回。認列口徑的全年 CapEx 見下方備忘。`],
                    [`② 租金（在帳＋表外）`, d.years.map(e => e.lease), void 0, void 0, void 0, `FY26 欄只含下半年；1H 租金 0.748 已含在 CFO 內。`],
                    [`③ 利息（含瀑布新債）`, d.years.map(e => e.interest), void 0, void 0, void 0, `FY26 欄只含下半年；1H 利息已含在 CFO 內。`],
                    [`④ JV／策略投資`, d.years.map((e, t) => t === 0 ? e.fyDiv : e.div), void 0, void 0, void 0, `FY26＝1H 實際 0.688 ＋ 下半年 1.25（已承諾餘額 1.15＋後續 0.1）。`],
                    [`⑤ 排程還本`, d.years.map((e, t) => t === 0 ? e.fyDebtPay : e.debtPay), void 0, void 0, void 0, `FY26＝1H 實際還款 5.219 ＋ 下半年到期表 4.413。`],
                    [`⑥ capped call（1H 實際）`, d.years.map((e, t) => t === 0 ? e.fyCapped : 0)],
                    [`用途合計`, d.years.map((e, t) => t === 0 ? e.fyUsesCash : e.uses), void 0, void 0, `tot`],
                    [`期前融資瀑布`, null],
                    [`融資前現金（扣既有新債利息）`, d.years.map(e => e.preCash), void 0, void 0, void 0, `＝期初現金 ＋ 營運來源 ＋ 可轉債／ATM − 用途（不含本期舉借利息）。`],
                    [`融資需求（補足至最低現金）`, d.years.map(e => e.need), void 0, void 0, `calc`, `＝MAX(0, 最低現金 − 融資前現金)。`],
                    [`期末 backlog`, d.years.map(e => e.backlogEnd), void 0, void 0, void 0, `＝期初 − 排程認列 − 新簽約認列 ＋ 新簽約增量 × 合約年期。期初 128.7（6/30 RPO 103.7＋Q3 新增 25）。`],
                    [`債務上限（債務／backlog）`, d.years.map(e => e.debtCap)],
                    [`期末總債務（既有＋可轉債＋新債）`, d.years.map(e => e.totalDebtEnd), void 0, void 0, `tot`],
                    [`每年股權吸收上限`, d.years.map(e => Number.isFinite(e.eqCap) ? e.eqCap : NaN), void 0, void 0, void 0, `＝現市值 × 上限 % × 期間長度。`],
                    [`新發行股數`, d.years.map(e => e.newShares), void 0, void 0, void 0, `＝股權募資 ÷ 發行價。`, `bn 股`, 3],
                    [`高息債餘額`, d.years.map(e => e.junkEnd), void 0, void 0, void 0],
                    [`融資前累積現金`, d.years.map(e => e.preFinCum), void 0, void 0, `tot`, `若不做任何新融資的累積現金；負值＝外部資金需求。`],
                    [`累計新股`, d.years.map(e => e.cumNewShares), void 0, void 0, `tot`, void 0, `bn 股`, 3],
                    [`缺口與現金`, null],
                    [`營運缺口（不含融資與還本）`, d.years.map((e, t) => t === 0 ? e.fyOperatingGap : e.operatingGap), void 0, void 0, `tot`, `本業能不能自給。FY26 的 1H 部分＝CFO 3.663 − 現金 CapEx 14.117 − JV 0.688（皆為 10-Q 實際值）。`],
                    [`期初累積現金`, d.years.map((e, t) => t === 0 ? ACTUAL_1H.cash1231 : d.years[t - 1].cum), void 0, void 0, void 0, `FY26 自 2025-12-31 的 3.127 起算。`],
                    [`1H 其他／受限現金調節`, d.years.map((e, t) => t === 0 ? e.hPlug : 0), void 0, void 0, void 0, `使 1H 實際流量接回 6/30 現金餘額；差額來自受限現金重分類與未逐項列出的項目。`],
                    [`期末累積現金`, d.years.map(e => e.cum), void 0, void 0, `tot`, `負值＝尚需向 DDTL／票據／股權市場籌措的金額。可用額度 10.0bn 未預先扣減。`],
                    [`FY26 全年備忘（認列口徑，對照公司指引）`, null],
                    [`CapEx 認列（1H 16.139＋下半年）`, d.years.map((e, t) => t === 0 ? e.fyGross : e.gross), void 0, void 0, void 0, `公司全年指引 35–39。`],
                    [`利息（1H 1.176＋下半年）`, d.years.map((e, t) => t === 0 ? e.fyInterest : e.interest), void 0, void 0, void 0, `Q3 指引 0.86–0.94／季。`],
                    [`租金現金（1H 0.748＋下半年）`, d.years.map((e, t) => t === 0 ? e.fyLease : e.lease)]
                  ]
                }), (0, $.jsxs)(`p`, {
                  className: `text-xs leading-relaxed text-muted`,
                  children: [`FY26 現金橋：2025-12-31 現金 `, Y(3.127), ` ＋ 1H 營運 `, Y(d.totals.h1Op), ` ＋ 1H 融資 `, Y(d.totals.h1Fin), ` − 1H 還款 `, Y(5.219), ` ＋ 其他／受限現金 `, Y(d.totals.h1Plug), ` ＝ 6/30 現金 `, Y(e.cash), `；再加 2H 缺口 `, Y(d.years[0].gap), ` ＝ FY26 期末 `, Y(d.years[0].cum), `。`]
                })]
              }), n === 2 && (0, $.jsxs)(`div`, {
                className: `space-y-3`,
                children: [(0, $.jsx)(hdrQ, {
                  title: `收入／產能輸入`,
                  tip: `收入 ＝ 排程 RPO − MAX(0, 排程 − 容量上限)；多出來的容量再乘「新產能簽約率」成為新簽約收入。rev/MW 錨點：期末 ARR 18.5–19.5 ÷ 1.85 GW ≈ 10.0–10.5 M/MW；7 月漲價約 25%、9/16 公告 Q3 以更高價簽短天期約，皆只影響新約。利用率已內含系統可用率，不另設 goodput 係數以免重複扣除。FY26 欄的 MW 為年底存量、收入為 2H 半年金額。`
                }), (0, $.jsx)(BM, {
                  rows: [
                    [`Accepted MW（期末主動電力）`, d.m.accepted, (e, t) => D(`accepted`, e, t), void 0, void 0, `YE26 指引 >1,850 MW；2030 目標 ≥8,000 MW。引擎會強制單調不減。`, `MW`],
                    [`Billable MW`, d.m.billable, (e, t) => D(`billable`, e, t), void 0, void 0, `引擎會強制不超過 Accepted。`, `MW`],
                    [`利用率`, e.m.util, (e, t) => D(`util`, e, t), void 0, void 0, `法說稱「近期產能實質售罄」，本模型不擬合為 100%。`, `%`],
                    [`每 MW 年收入`, e.m.revMW, (e, t) => D(`revMW`, e, t), 5e-4, void 0, `期末 ARR 18.5–19.5 ÷ 1.85 GW ≈ 10.0–10.5 M/MW。`, `US$bn/MW`],
                    [`新產能簽約率`, e.m.fill, (e, t) => D(`fill`, e, t), 1, void 0, `把這欄調成 0，就能看到只靠期初 RPO 的缺口有多大——最重要的壓力測試。`, `%`],

                  ]
                })]
              }), n === 3 && (0, $.jsxs)(`div`, {
                className: `space-y-3`,
                children: [(0, $.jsx)(hdrQ, {
                  title: `電力成本（overlay，預設關閉）`,
                  tip: `電力 ＝ Accepted MW × 8760 × PUE × 電價 ÷ 1e9 × 期間係數（十億美元）；維護 ＝ Accepted × 單價 ÷ 1000 × 期間係數。CoreWeave 的電費多由房東轉嫁、已含在營收成本內，因此預設不另扣；若貢獻率改用未扣電費的毛數，才打開這個 overlay。`
                }), (0, $.jsx)(BM, {
                  rows: [
                    [`Accepted MW`, d.m.accepted, (e, t) => D(`accepted`, e, t), void 0, void 0, void 0, `MW`],
                    [`電價`, e.m.power, (e, t) => D(`power`, e, t), void 0, void 0, void 0, `$/MWh`],
                    [`PUE`, e.m.pue, (e, t) => D(`pue`, e, t), .01, void 0, void 0, ``],
                    [`維護`, e.m.maint, (e, t) => D(`maint`, e, t), .01, void 0, void 0, `US$m/MW`],
                    [`電力（計算）`, d.years.map(e => e.power), void 0, void 0, `calc`],
                    [`維護（計算）`, d.years.map(e => e.maint), void 0, void 0, `calc`]
                  ]
                })]
              }), n === 4 && (0, $.jsxs)(`div`, {
                className: `space-y-3`,
                children: [(0, $.jsx)(hdrQ, {
                  title: `信用與利率`,
                  tip: `信用損失 ＝ 收入 × 違約率 × (1−回收率)。前三大客戶占 72% 營收、Jane Street 為私人公司，違約率是對此集中度的定價，不是預測。新債利率是瀑布「新增借款」的成本（2031 票據 9.75%、2032 9.625%、DDTL 5.0 SOFR+450），不是加權平均 7.8%；存量債務的利息在「收支假設」頁單獨輸入。`
                }), (0, $.jsx)(BM, {
                  rows: [
                    [`客戶違約率`, e.m.defaultP, (e, t) => D(`defaultP`, e, t), void 0, void 0, void 0, `%`],
                    [`回收率`, e.m.recovery, (e, t) => D(`recovery`, e, t), void 0, void 0, void 0, `%`],
                    [`新債利率`, e.m.rate, (e, t) => D(`rate`, e, t), void 0, void 0, `瀑布新債（額度＋資產層）的利率。CDS 開關打開後會加上 MAX(0, CDS−450)×0.4bp。`, `%`],
                    [`客戶預付`, e.a.customerFund.map(e => e * 100), (e, t) => E(`customerFund`, e, t / 100), void 0, void 0, void 0, `%`]
                  ]
                })]
              }), n === 5 && (0, $.jsx)(`table`, {
                className: `w-full min-w-[480px] text-sm`,
                children: (0, $.jsxs)(`tbody`, {
                  children: [
                    [
                      [`資產殘值 %（GPU 五期毛 CapEx）`, `residual`],
                      [`再出租率 %`, `rerent`],
                      [`剩餘 RPO 現金率 %`, `margin`],
                      [`租賃剩餘年數`, `residualLeaseYears`]
                    ].map(([n, r]) => (0, $.jsxs)(`tr`, {
                      className: `border-b border-border`,
                      children: [(0, $.jsx)(`td`, {
                        className: `py-2 font-medium`,
                        children: n
                      }), (0, $.jsx)(`td`, {
                        className: `py-2`,
                        children: (0, $.jsx)(IM, {
                          value: e.terminal[r],
                          onChange: e => t(t => ({
                            ...t,
                            terminal: {
                              ...t.terminal,
                              [r]: e
                            },
                            scenario: `custom`
                          }))
                        })
                      })]
                    }, r)), (0, $.jsxs)(`tr`, {
                      className: `border-b border-border`,
                      children: [(0, $.jsx)(`td`, {
                        className: `py-2 font-medium`,
                        children: `連動租賃尾端`
                      }), (0, $.jsxs)(`td`, {
                        className: `py-2 font-mono tabular-nums`,
                        children: [Y(d.leaseTail), `bn =（在帳 FY30 `, LEASE_CASH_ON_BAL[4], ` + 表外現金租金 `, e.a.newLease[4], `）×`, ` `, e.terminal.residualLeaseYears, ` 年（10-Q 加權剩餘租期 12 年）。五期表外現金 `, Y(S.cashFive), ` + 尾 `, Y(S.tail), ` = `, Y(S.mapped), ` vs 承諾 50.2；在帳 2030 後尚有 `, Y(LEASE_AFTER_FY30), `。`]
                      })]
                    }), (0, $.jsxs)(`tr`, {
                      children: [(0, $.jsx)(`td`, {
                        className: `py-2 font-medium`,
                        children: `終值淨額`
                      }), (0, $.jsxs)(`td`, {
                        className: `py-2 font-mono tabular-nums`,
                        children: [Y(d.totals.terminal), `bn（法說：A100 續約至 2029、舊世代仍售罄——殘值假設 25% 偏保守，但 GPU 折舊爭議未解）`]
                      })]
                    })
                  ]
                })
              }), n === 6 && (0, $.jsxs)(`div`, {
                className: `space-y-3`,
                children: [(0, $.jsx)(hdrQ, {
                  title: `站點`,
                  tip: `具名站點可編輯。公司揭露 51 站、1.5 GW 主動、4.2 GW 簽約，但不給逐站 MW；下列 MW 多為房東公告或推定（Analogy／Assumed）。最後一列是殘差＝年度 Accepted − 具名加總；殘差偏大是必然，代表容量預測多數不是由具名站點支撐。`
                }), (0, $.jsxs)(`table`, {
                  className: `w-full min-w-[960px] text-xs`,
                  children: [(0, $.jsx)(`thead`, {
                    children: (0, $.jsx)(`tr`, {
                      className: `text-muted`,
                      children: [`站點`, `房東／型態`, `Planned`, `Energized`, `Accepted`, `Billable`, `狀態`, `里程碑`, `日期`, `信心`].map(e => (0, $.jsx)(`th`, {
                        className: `px-1 py-2 text-left font-medium`,
                        children: e
                      }, e))
                    })
                  }), (0, $.jsx)(`tbody`, {
                    children: d.sites.map((e, t) => (0, $.jsx)(`tr`, {
                      className: `border-t border-border`,
                      children: [`name`, `operator`, `planned`, `energized`, `accepted`, `billable`, `status`, `next`, `date`, `confidence`].map(n => (0, $.jsx)(`td`, {
                        className: `px-1 py-1`,
                        children: e.residual || ![`planned`, `energized`, `accepted`, `billable`, `name`, `operator`, `status`, `next`, `date`, `confidence`].includes(n) ? (0, $.jsx)(`span`, {
                          className: Vj(`block px-1 py-2`, e.residual && `text-muted`),
                          children: String(e[n])
                        }) : typeof e[n] == `number` ? (0, $.jsx)(IM, {
                          value: e[n],
                          onChange: e => O(t, n, e),
                          step: 1
                        }) : (0, $.jsx)(`input`, {
                          value: String(e[n]),
                          onChange: e => O(t, n, e.target.value),
                          className: `h-9 w-full rounded-sm border border-border bg-card px-2 text-xs`
                        })
                      }, n))
                    }, e.id))
                  })]
                }), (0, $.jsx)(hdrQ, {
                  title: `站點租賃：房東揭露 vs 10-Q 總額`,
                  tip: `10-Q 不揭露逐站租金；下表為房東公告的合約值（Interested-party），用來檢驗每 MW 年租金與模型租金路徑是否一致。10-Q「單站 393 MW、上限 14.7／16 年」＝Helios Phase II 260＋Phase III 133，MW 完全吻合；「按造價計租 355 MW」推定含 Kenilworth（10-Q 明寫該 JV 租金依建造成本計），逐站歸屬未揭露。`
                }), (0, $.jsx)(`div`, {
                  style: xstyQ.wrap,
                  children: (0, $.jsxs)(`table`, {
                    style: {
                      ...xstyQ.table,
                      minWidth: 860
                    },
                    children: [(0, $.jsx)(`thead`, {
                      children: (0, $.jsx)(`tr`, {
                        children: [`站點`, `契約 MW`, `合約總值 $bn`, `年期`, `年租金 $bn`, `每 MW 年租金 $m`, `10-Q 對應`].map((e, t) => (0, $.jsx)(`th`, {
                          style: t === 0 ? xstyQ.thL : xstyQ.th,
                          children: e
                        }, e))
                      })
                    }), (0, $.jsxs)(`tbody`, {
                      children: [...e.sites.filter(e => e.contract).map((e, t) => (0, $.jsxs)(`tr`, {
                        children: [(0, $.jsx)(`td`, {
                          style: {
                            ...xstyQ.tdL,
                            background: t % 2 ? `#fbfaf7` : `#fff`
                          },
                          children: e.name
                        }), [e.planned, e.contract, e.years, e.contract / e.years, e.contract / e.years / e.planned * 1e3].map((e, n) => (0, $.jsx)(`td`, {
                          style: {
                            ...xstyQ.td,
                            background: t % 2 ? `#fbfaf7` : `#fff`
                          },
                          children: n === 0 || n === 2 ? Y(e, 0) : Y(e, 2)
                        }, n)), (0, $.jsx)(`td`, {
                          style: {
                            ...xstyQ.td,
                            textAlign: `left`,
                            fontFamily: `ui-sans-serif, system-ui, sans-serif`,
                            fontSize: 11,
                            color: `#6b7280`,
                            background: t % 2 ? `#fbfaf7` : `#fff`
                          },
                          children: e.id === `helios` ? `未交付 393 MW＝單站上限 14.7` : e.id === `polaris` ? `大致在未起租 35.5 內` : `部分在帳、部分未起租`
                        })]
                      }, e.id)), (0, $.jsxs)(`tr`, {
                        children: [(0, $.jsx)(`td`, {
                          style: {
                            ...xstyQ.tdL,
                            background: `#eef2f7`,
                            fontWeight: 700
                          },
                          children: `具名合計／加權`
                        }), [e.sites.filter(e => e.contract).reduce((e, t) => e + t.planned, 0), e.sites.filter(e => e.contract).reduce((e, t) => e + t.contract, 0), NaN, e.sites.filter(e => e.contract).reduce((e, t) => e + t.contract / t.years, 0), d.totals.bench].map((e, t) => (0, $.jsx)(`td`, {
                          style: {
                            ...xstyQ.td,
                            background: `#eef2f7`,
                            fontWeight: 700
                          },
                          children: Number.isFinite(e) ? t === 0 ? Y(e, 0) : Y(e, 2) : ``
                        }, t)), (0, $.jsx)(`td`, {
                          style: {
                            ...xstyQ.td,
                            background: `#eef2f7`
                          },
                          children: ``
                        })]
                      })]
                    })]
                  })
                }), (0, $.jsx)(BM, {
                  rows: [
                    [`10-Q 對帳`, null],
                    [`10-Q 已揭露租賃承諾（在帳 29.4＋未起租 35.5＋單站上限 14.7）`, [LEASE_FACTS.onBal + LEASE_FACTS.notCommenced + LEASE_FACTS.singleCap, 0, 0, 0, 0], void 0, void 0, void 0, `另有 355 MW 按造價計租，金額未定，未含。`],
                    [`減：具名站點合約總值`, [-e.sites.filter(e => e.contract).reduce((e, t) => e + t.contract, 0), 0, 0, 0, 0]],
                    [`＝ 未具名站點（殘差）`, [LEASE_FACTS.onBal + LEASE_FACTS.notCommenced + LEASE_FACTS.singleCap - e.sites.filter(e => e.contract).reduce((e, t) => e + t.contract, 0), 0, 0, 0, 0], void 0, void 0, `tot`, `43 座以上資料中心中未具名者（託管商與其他房東）。殘差占比高是必然——公司不揭露逐站。`],
                    [`租金路徑檢驗（第三方租賃占 85%）`, null],
                    [`模型租金（在帳＋表外）`, d.years.map(e => e.lease)],
                    [`模型每 MW 年租金`, d.years.map(e => e.rentPerMW), void 0, void 0, `calc`, `＝模型租金 ÷ 期間長度 ÷ 平均 Accepted MW。FY26 欄只含下半年租金、MW 取全年平均，略低估。`, `US$m/MW`],
                    [`基準租金（MW × 市場基準 × 85%）`, d.years.map(e => e.rentBench), void 0, void 0, void 0, `市場基準＝具名站點加權每 MW 年租金。若 8 GW 多數比照 Helios／APLD／CORZ 的租金水準。`],
                    [`差額（基準 − 模型）`, d.years.map(e => e.rentBench - e.lease), void 0, void 0, `tot`, `正值＝模型租金可能低估的金額。本版只作檢驗，未改為 MW 驅動，以免與 CapEx 同時放大同一條 8 GW 假設。`]
                  ]
                })]
              }), n === 7 && (0, $.jsxs)(`div`, {
                className: `space-y-3`,
                children: [(0, $.jsx)(hdrQ, {
                  title: `敏感性：加權目標價（基準 $${Y(m[0]?.base ?? 0, 1)}）`,
                  tip: `每一列只改動一個變數、其餘維持目前輸入，重新跑完整的資金瀑布與評價。長條為相對基準的變動（美元／股），依影響幅度排序。期前融資下期末現金固定在最低現金，所以改以目標價、股權需求與融資前缺口作為指標。每 MW 年收入會同時改變容量上限、RPO 可轉換收入、新簽約收入、現金貢獻、融資缺口與估值；每 MW 建置成本、GPU 壽命、FY31 新增 MW 等「做多少、花多少」的變數也會同時改變融資需求。`,
                  w: 460
                }), (0, $.jsx)(`div`, {
                  className: `h-[28rem]`,
                  style: {
                    height: 460
                  },
                  children: (0, $.jsx)(Gs, {
                    width: `100%`,
                    height: `100%`,
                    children: (0, $.jsxs)(Lk, {
                      layout: `vertical`,
                      data: m.map(e => ({
                        name: e.name,
                        下行: Number((e.low - e.base).toFixed(1)),
                        上行: Number((e.high - e.base).toFixed(1))
                      })),
                      stackOffset: `sign`,
                      margin: {
                        left: 24,
                        right: 16
                      },
                      children: [(0, $.jsx)(uD, {
                        stroke: `#ddd6c8`,
                        strokeDasharray: `3 3`
                      }), (0, $.jsx)(XD, {
                        type: `number`,
                        tick: {
                          fontSize: 11
                        }
                      }), (0, $.jsx)(pO, {
                        type: `category`,
                        dataKey: `name`,
                        width: 130,
                        tick: {
                          fontSize: 11
                        }
                      }), (0, $.jsx)(Es, {}), (0, $.jsx)(ow, {
                        dataKey: `下行`,
                        stackId: `t`,
                        fill: `#9f1239`
                      }), (0, $.jsx)(ow, {
                        dataKey: `上行`,
                        stackId: `t`,
                        fill: `#0f6b4c`
                      })]
                    })
                  })
                }), (0, $.jsx)(`div`, {
                  style: xstyQ.wrap,
                  children: (0, $.jsxs)(`table`, {
                    style: {
                      ...xstyQ.table,
                      minWidth: 900
                    },
                    children: [(0, $.jsx)(`thead`, {
                      children: (0, $.jsx)(`tr`, {
                        children: [`變數`, `設定 A`, `設定 B`, `目標價 A`, `目標價 B`, `股權需求 A`, `股權需求 B`, `融資前缺口 A`, `融資前缺口 B`].map((e, t) => (0, $.jsx)(`th`, {
                          style: t === 0 ? xstyQ.thL : xstyQ.th,
                          children: e
                        }, e))
                      })
                    }), (0, $.jsxs)(`tbody`, {
                      children: [(0, $.jsx)(`tr`, {
                        children: [`基準（目前輸入）`, `—`, `—`, `$${Y(m[0]?.base ?? 0,1)}`, ``, Y(m[0]?.baseEq ?? 0,1), ``, Y(m[0]?.baseGap ?? 0,1), ``].map((e, t) => (0, $.jsx)(`td`, {
                          style: {
                            ...(t === 0 ? xstyQ.tdL : xstyQ.td),
                            background: `#eef2f7`,
                            fontWeight: 700
                          },
                          children: e
                        }, t))
                      }), ...m.map((e, t) => (0, $.jsx)(`tr`, {
                        children: [e.name, e.la, e.lb, `$${Y(e.a.tgt,1)}`, `$${Y(e.b.tgt,1)}`, Y(e.a.eq,1), Y(e.b.eq,1), Y(e.a.gap,1), Y(e.b.gap,1)].map((n, r) => (0, $.jsx)(`td`, {
                          style: {
                            ...(r === 0 ? xstyQ.tdL : xstyQ.td),
                            background: t % 2 ? `#fbfaf7` : `#fff`,
                            color: r === 3 && e.a.tgt < e.base - .05 || r === 4 && e.b.tgt < e.base - .05 ? `#9f1239` : r === 3 && e.a.tgt > e.base + .05 || r === 4 && e.b.tgt > e.base + .05 ? `#0f6b4c` : `#1f2937`
                          },
                          children: n
                        }, r))
                      }, e.name))]
                    })]
                  })
                }), (0, $.jsx)(`p`, {
                  className: `text-xs text-muted`,
                  children: `股權需求與融資前缺口單位為 US$bn（模型期合計）。「每 MW 年收入」敏感性將 FY26–FY30 五期 Revenue/MW 同比調整 ±15%，其他假設不變，並重新計算容量收入上限、資金瀑布及評價。WACC 與 EV/EBITDA 倍數只影響評價，不改變融資。`
                })]
              }), n === 8 && (0, $.jsxs)(`ul`, {
                className: `space-y-2`,
                children: [d.checks.map(e => (0, $.jsxs)(`li`, {
                  className: `rounded-md border border-border bg-surface px-3 py-2`,
                  children: [(0, $.jsxs)(`div`, {
                    className: `flex items-center gap-2`,
                    children: [(0, $.jsx)(Uj, {
                      tone: e.ok ? e.severity === `watch` ? `watch` : `ok` : e.severity === `block` ? `bad` : `watch`,
                      children: e.ok ? e.severity === `watch` ? `觀察` : `通過` : `不一致`
                    }), (0, $.jsx)(`span`, {
                      className: `text-sm font-medium`,
                      children: e.title
                    })]
                  }), (0, $.jsx)(`p`, {
                    className: `mt-1 text-xs leading-relaxed text-muted`,
                    children: e.detail
                  })]
                }, e.id)), p.map(e => (0, $.jsxs)(`li`, {
                  className: `rounded-md border border-border bg-surface px-3 py-2`,
                  children: [(0, $.jsxs)(`div`, {
                    className: `flex items-center gap-2`,
                    children: [(0, $.jsx)(Uj, {
                      tone: e.ok ? e.severity === `watch` ? `watch` : `ok` : `bad`,
                      children: e.ok ? e.severity === `watch` ? `觀察` : `通過` : `不一致`
                    }), (0, $.jsx)(`span`, {
                      className: `text-sm font-medium`,
                      children: e.title
                    })]
                  }), (0, $.jsx)(`p`, {
                    className: `mt-1 text-xs leading-relaxed text-muted`,
                    children: e.detail
                  })]
                }, e.id))]
              }), n === 13 && (0, $.jsx)(QuarterTabQ, { d: d, p: f, st: e }), n === 14 && (0, $.jsx)(PerMwTabQ, { d: d, st: e, o: o }), n === 12 && (0, $.jsxs)(`div`, {
                className: `space-y-3`,
                children: [(0, $.jsx)(hdrQ, {
                  title: `版本紀錄（基準情境目標價變化與原因）`,
                  tip: `每一版的主要變更、基準情境加權目標價與變動原因。v2.9 以前的檔名日期統一沿用 20260918，此處日期為實際修改日。v1.7 以前的基準為 8 GW、v1.7 為 5 GW、v1.8 起為 5.6 GW，情境定義改變的版本不可直接比較。目標價以當版預設輸入計算（DCF 0 截斷）。`,
                  w: 460
                }), (0, $.jsx)(`div`, {
                  style: xstyQ.wrap,
                  children: (0, $.jsxs)(`table`, {
                    style: {
                      ...xstyQ.table,
                      minWidth: 900
                    },
                    children: [(0, $.jsx)(`thead`, {
                      children: (0, $.jsx)(`tr`, {
                        children: [`版本`, `日期`, `主要變更`, `基準目標價`, `變動原因`].map((e, t) => (0, $.jsx)(`th`, {
                          style: t === 0 ? xstyQ.thL : t === 2 || t === 4 ? {
                            ...xstyQ.th,
                            textAlign: `left`
                          } : xstyQ.th,
                          children: e
                        }, e))
                      })
                    }), (0, $.jsx)(`tbody`, {
                      children: [...VLOG].reverse().map((e, t) => (0, $.jsx)(`tr`, {
                        children: e.map((e, n) => (0, $.jsx)(`td`, {
                          style: {
                            ...(n === 0 ? xstyQ.tdL : xstyQ.td),
                            background: t % 2 ? `#fbfaf7` : `#fff`,
                            ...(n === 2 || n === 4 ? {
                              textAlign: `left`,
                              whiteSpace: `normal`,
                              fontFamily: `ui-sans-serif, system-ui, sans-serif`,
                              fontSize: 11.5,
                              minWidth: n === 2 ? 420 : 200
                            } : {}),
                            fontWeight: n === 3 ? 700 : 400
                          },
                          children: e
                        }, n))
                      }, e[0]))
                    })]
                  })
                })]
              }), n === 11 && (0, $.jsxs)(`div`, {
                className: `space-y-3`,
                children: [(0, $.jsx)(hdrQ, {
                  title: `新債與新股（類資產負債表：現金、債務、股本）`,
                  tip: `把期前融資瀑布的結果按期整理成存量：期初現金＋營運收支淨額＋瀑布融資−新融資利息＝期末現金；既有債務依 10-Q 到期表遞減，加上 9/17 可轉債、瀑布新債與高息債＝總債務。股數＝基礎股數（含 ATM 上限、每年 SBC 稀釋 1%）＋累計新股。FY26 欄的現金自 6/30 起算（模型期）。`,
                  w: 480
                }), (0, $.jsx)(BM, {
                  rows: [
                    [`現金`, null],
                    [`期初現金`, d.years.map((t, n) => n === 0 ? e.cash : d.years[n - 1].cum), void 0, void 0, void 0, `FY26 為 2026-06-30 現金（模型期起點）。`],
                    [`營運收支淨額（含 9/17 可轉債／ATM，不含瀑布）`, d.years.map(e => e.preFinGap), void 0, void 0, void 0, `＝營運來源＋可轉債／ATM − 所有支出（含還本、不含新融資利息）。負值即當期外部資金需求。`],
                    [`＋ 瀑布：新債（額度＋資產層）`, d.years.map(e => e.newDebt)],
                    [`＋ 瀑布：高息債`, d.years.map(e => e.junk)],
                    [`＋ 瀑布：股權募資`, d.years.map(e => e.equity)],
                    [`− 新融資利息（新債＋高息債）`, d.years.map(e => -e.newDebtInt)],
                    [`期末現金`, d.years.map(e => e.cum), void 0, void 0, `tot`, `期前融資使期末現金不低於最低現金。`],
                    [`債務`, null],
                    [`既有債務期初（10-Q 本金）`, d.years.map((t, n) => n === 0 ? (e.includeDebt ? DBT_P : DBT_P) : d.years[n - 1].existDebtEnd - 3.7)],
                    [`− 排程還本`, d.years.map(e => -e.debtPay)],
                    [`既有債務期末`, d.years.map(e => e.existDebtEnd - 3.7), void 0, void 0, `calc`],
                    [`＋ 9/17 可轉債（2033 到期）`, d.years.map(() => 3.7)],
                    [`＋ 瀑布新債餘額`, d.years.map(e => e.newDebtEnd)],
                    [`＋ 高息債餘額`, d.years.map(e => e.junkEnd)],
                    [`總債務`, d.years.map(e => e.totalDebtEnd), void 0, void 0, `tot`],
                    [`淨負債（總債務 − 期末現金）`, d.years.map(e => e.totalDebtEnd - e.cum), void 0, void 0, `tot`],
                    [`股本`, null],
                    [`基礎股數（含 ATM 上限、SBC 稀釋）`, f.fwd.map((e, t) => e.shares - d.years[t].cumNewShares), void 0, void 0, void 0, void 0, `bn 股`, 3],
                    [`＋ 本期新股`, d.years.map(e => e.newShares), void 0, void 0, void 0, `＝股權募資 ÷ 發行價。`, `bn 股`, 3],
                    [`總股數（期末）`, f.fwd.map(e => e.shares), void 0, void 0, `tot`, void 0, `bn 股`, 3],
                    [`槓桿`, null],
                    [`期末 backlog`, d.years.map(e => e.backlogEnd)],
                    [`債務上限（債務／backlog × 期末 backlog）`, d.years.map(e => e.debtCap)],
                    [`總債務 ÷ 期末 backlog`, d.years.map(e => e.totalDebtEnd / Math.max(e.backlogEnd, .01)), void 0, void 0, void 0, `高息債不受 backlog 上限約束，所以此比率可能超過上限。`, `x`, 2],
                    [`總債務 ÷ EBITDA（年化）`, d.years.map((e, t) => e.totalDebtEnd / Math.max(e.ebitdaPL / PERIOD_YEARS[t], .01)), void 0, void 0, void 0, `FY26 模型期 EBITDA 以半年 ×2 年化。`, `x`, 1]
                  ]
                })]
              }), n === 10 && (0, $.jsxs)(`div`, {
                className: `space-y-3`,
                children: [(0, $.jsx)(hdrQ, {
                  title: `既有債務（10-Q Note 10，2026-06-30）`,
                  tip: `有效利率為 10-Q 揭露值（含折價攤銷，四捨五入至整數 %）。存量利息＝平均本金（依 10-Q 到期表遞減）× 加權有效利率 × 期間長度，另加 9/17 可轉債利息與 FY26 校準列（使下半年對上 Q3 指引 0.86–0.94／季）。限制：假設利率組合不變——15% 的 DDTL 1.0 於 2028 年先到期，後期實際利率會略低。公司簡報的 7.8% 為票面口徑、含 1.75% 可轉債稀釋效果，與此處有效利率 8.4% 口徑不同。`
                }), (0, $.jsx)(`div`, {
                  style: xstyQ.wrap,
                  children: (0, $.jsxs)(`table`, {
                    style: {
                      ...xstyQ.table,
                      minWidth: 820
                    },
                    children: [(0, $.jsx)(`thead`, {
                      children: (0, $.jsx)(`tr`, {
                        children: [`工具`, `類別`, `到期`, `有效利率`, `6/30 本金`, `年化利息`, `備註`].map((e, t) => (0, $.jsx)(`th`, {
                          style: t === 0 ? xstyQ.thL : xstyQ.th,
                          children: e
                        }, e))
                      })
                    }), (0, $.jsxs)(`tbody`, {
                      children: [...DEBT_TOOLS.map((e, t) => (0, $.jsx)(`tr`, {
                        children: [e[0], e[1], e[2], hA(e[3] * 100, 0), Y(e[4], 3), Y(e[3] * e[4], 3), e[5]].map((e, n) => (0, $.jsx)(`td`, {
                          style: {
                            ...(n === 0 ? xstyQ.tdL : xstyQ.td),
                            background: t % 2 ? `#fbfaf7` : `#fff`,
                            ...(n === 1 || n === 2 || n === 6 ? {
                              textAlign: `left`,
                              fontSize: 11,
                              color: `#6b7280`,
                              fontFamily: `ui-sans-serif, system-ui, sans-serif`
                            } : {})
                          },
                          children: e
                        }, n))
                      }, e[0])), (0, $.jsx)(`tr`, {
                        children: [`合計／加權`, ``, ``, hA(DBT_R * 100, 1), Y(DBT_P, 3), Y(DBT_I, 3), `＝10-Q 本金 35.551`].map((e, t) => (0, $.jsx)(`td`, {
                          style: {
                            ...(t === 0 ? xstyQ.tdL : xstyQ.td),
                            background: `#eef2f7`,
                            fontWeight: 700
                          },
                          children: e
                        }, t))
                      }), (0, $.jsx)(`tr`, {
                        children: [`9/17 期後：2033 可轉債`, `追索`, `2033`, `2.875%`, Y(3.7, 3), Y(CONV_I, 3), `轉換價 $97.85；不在五期還本表`].map((e, t) => (0, $.jsx)(`td`, {
                          style: {
                            ...(t === 0 ? xstyQ.tdL : xstyQ.td),
                            background: `#fff`
                          },
                          children: e
                        }, t))
                      })]
                    })]
                  })
                }), (0, $.jsx)(BM, {
                  rows: [
                    [`存量利息推導（本金依 10-Q 到期表遞減）`, null],
                    [`期初本金`, d.years.map(e => e.pBeg)],
                    [`排程還本`, [...DEBT_AMORT]],
                    [`期末本金`, d.years.map(e => e.pEnd)],
                    [`存量債務利息（含可轉債與 FY26 校準）`, d.years.map(e => e.intStock), void 0, void 0, `tot`, `＝(期初＋期末)÷2 × 8.4% × 期間長度 ＋ 3.7×2.875%×期間長度 ＋ FY26 校準 0.35。`],
                    [`對照：v1.4 手動值`, d.years.map(e => e.intOld)],
                    [`瀑布新債利息（計算）`, d.years.map(e => e.newDebtInt), void 0, void 0, `calc`, `新債利率 × 新債餘額（含本期舉借）。`]
                  ]
                })]
              }), n === 9 && (0, $.jsxs)(`div`, {
                className: `space-y-4 text-sm leading-relaxed`,
                children: [(0, $.jsx)(`p`, {
                  className: `text-xs leading-relaxed text-muted`,
                  children: SOURCE_ORDER_NOTE
                }), (0, $.jsx)(`div`, {
                  className: `overflow-x-auto`,
                  children: (0, $.jsxs)(`table`, {
                    className: `w-full min-w-[720px] text-xs`,
                    children: [(0, $.jsx)(`thead`, {
                      children: (0, $.jsxs)(`tr`, {
                        className: `border-b border-border text-left text-muted`,
                        children: [(0, $.jsx)(`th`, {
                          className: `py-2 pr-3 font-medium`,
                          children: `項目`
                        }), (0, $.jsx)(`th`, {
                          className: `py-2 pr-3 font-medium`,
                          children: `法說／8-K（前瞻）`
                        }), (0, $.jsx)(`th`, {
                          className: `py-2 pr-3 font-medium`,
                          children: `10-Q（入帳）`
                        }), (0, $.jsx)(`th`, {
                          className: `py-2 font-medium`,
                          children: `本模型取捨`
                        })]
                      })
                    }), (0, $.jsx)(`tbody`, {
                      className: `font-mono tabular-nums`,
                      children: [
                        [`RPO／backlog`, `backlog 104；Q3 初 +>25；>50% 已開始交付`, `RPO 103.7 · 41／39／20`, `RPO 聽 10-Q；25 另列、按同形狀分攤（Derived）`],
                        [`期初現金`, `6.9bn「含受限現金」`, `現金 5.524＋有價證券 0.015；受限 1.38`, `期初用 5.5，受限現金不計`],
                        [`2026 CapEx`, `35–39（上修）；Q3 11.5–13.5`, `H1 認列 16.139；現金購置 14.117；CIP 11.9`, `v1.5 改由 MW 推導：YE25 850→YE26 1,850 MW、λ 35%、$34m/MW → 全年 37.0；FY27 起同一公式`],
                        [`客戶預付`, `「customer prepayments」為籌資來源之一`, `遞延收入 9.692；H1 淨流入 1.365（占 H1 CapEx 8.5%）`, `預付比率 8%；不降專案總成本`],
                        [`在帳租賃`, `未量化`, `未折現 29.135＋0.235；負債 16.319＋0.221；H1 現金 0.748`, `聽 10-Q 到期表，固定輸入`],
                        [`表外租賃`, `未提金額`, `未起租 35.5（2026–29 起租、7–16 年）＋單站 393 MW 上限 14.7＋355 MW 按造價`, `現金路徑五期 ${Y(S.cashFive,1)}＋尾 ${Y(S.tail,0)}，對 50.2 差 ${Y(S.gap,0)}`],
                        [`債務`, `「$32bn+ 資本已鎖定」；加權利率降近 300bps`, `本金 35.551；2H26–FY30 攤還 20.655；可用額度 10.014`, `攤還聽 10-Q（基本情境）；額度是來源不是現金，不預先計入`],
                        [`利息`, `Q3 0.86–0.94／季`, `Q2 費用 0.64；16 筆工具加權有效利率 8.4%`, `v1.5：存量利息＝遞減本金×8.4%（見既有債務）＋FY26 校準 0.35；新借款以瀑布新債利率 9% 計`],
                        [`9/17 籌資`, `可轉債 3.7（2.875%、$97.85）；ATM ≤35m 股`, `期後事項，10-Q 未含`, `淨現金 3.2（扣承銷與 capped call）＋ATM 上限 2.8 = 6.0 記於 2H26；股數 0.5865；可轉債利息入基本利息，本金 2033 到期不在五期還本表`],
                        [`主動／簽約電力`, `1.5 GW／4.2 GW；YE >1.85；2030 ≥8 GW`, `無 MW 揭露（僅 393＋355 MW 未交付）`, `Accepted 路徑對公司指引，不獨立驗證`],
                        [`利用率／售罄`, `近期產能實質售罄；A100 續約至 2029`, `無`, `基準 95%，不擬合 100%`],
                        [`定價`, `7 月 SKU 漲價約 25%；新約貢獻率 +5–10pt`, `無`, `rev/MW 2H26 0.0112→FY27 0.0115；貢獻率不因法說上調`],
                        [`營收／營益指引`, `12.4–13.2；調整後營益 0.96–1.15；Q4 低雙位數 %`, `H1 4.653；GAAP 營損 −0.193`, `2H26 IS 對上指引；GAAP 營利率不擬合調整後口徑`],
                        [`客戶集中`, `「10 家 >$1bn 承諾」`, `A 36%／B 26%／C 10%；應收 A、B 各 32%`, `違約率 0.5→2.5%、回收 50%`],
                        [`FCF 轉正`, `未給時點；「合約到期時叢集無槓桿」`, `H1 CFO 3.663 − 現金購置 14.117`, `不編時點。五期 UFCF 為負`],
                        [`5Y CDS`, `未提`, `非申報項目`, `中價 720bps、Bid/Ask 690/750、區間 680–855（User-provided，2026-09）。新債利率基準 9% 已含此水準，預設不再疊加傳導`],
                        [`未動用額度`, `「$32bn+ 資本已鎖定」`, `可用額度 10.014（RCF＋DDTL）`, `預設不計入現金；設開關可改採流動性口徑`],
                        [`Q3 定價`, `9/16 公告：6/30 後以更高價簽約、Q3 已簽短天期合約`, `期後事項`, `不上修 rev/MW；記為觀察（DDTL 5.5 支援短約是同一條敘事）`]
                      ].map(([e, t, n, r]) => (0, $.jsxs)(`tr`, {
                        className: `border-b border-border/70`,
                        children: [(0, $.jsx)(`td`, {
                          className: `py-1.5 pr-3 font-sans font-medium`,
                          children: e
                        }), (0, $.jsx)(`td`, {
                          className: `py-1.5 pr-3 text-muted`,
                          children: t
                        }), (0, $.jsx)(`td`, {
                          className: `py-1.5 pr-3`,
                          children: n
                        }), (0, $.jsx)(`td`, {
                          className: `py-1.5 font-sans text-muted`,
                          children: r
                        })]
                      }, e))
                    })]
                  })
                }), (0, $.jsxs)(Jj, {
                  className: `border-border bg-surface p-4`,
                  children: [(0, $.jsx)(`h3`, {
                    className: `font-display text-base font-semibold`,
                    children: `表外與契約性支出總表（10-Q，6/30/2026，US$bn）`
                  }), (0, $.jsx)(`div`, {
                    className: `mt-2 overflow-x-auto`,
                    children: (0, $.jsxs)(`table`, {
                      className: `w-full min-w-[720px] text-xs`,
                      children: [(0, $.jsx)(`thead`, {
                        children: (0, $.jsx)(`tr`, {
                          className: `border-b border-border text-muted`,
                          children: [`項目`, `金額`, `時程`, `入表？`, `模型位置`].map((e, t) => (0, $.jsx)(`th`, {
                            className: Vj(`py-2 pr-3 font-medium`, t === 1 ? `text-right` : `text-left`),
                            children: e
                          }, e))
                        })
                      }), (0, $.jsx)(`tbody`, {
                        children: [
                          [`債務本金（含 DDTL 13.6、票據 16.6、OEM 5.1、可轉債 6.6）`, `35.55`, `2026 餘 4.41／27 6.18／28 4.42／29 2.42／30 3.22／後 14.90`, `是`, `排程還本（固定）`],
                          [`在帳營業＋融資租賃（未折現）`, `29.37`, `2026 餘 1.04／27 2.34／28 2.31／29 2.37／30 2.29／後 19.02`, `是（PV 16.54）`, `在帳現金租金（固定）`],
                          [`未起租租賃（未折現）`, `35.5`, `2026–2029 起租，7–16 年`, `否`, `表外現金租金（假設路徑）`],
                          [`單站按造價計租，上限`, `14.7`, `393 MW 未交付，2026／2028 分期；16 年`, `否`, `表外現金租金（併入路徑）`],
                          [`按造價計租，355 MW`, `未定`, `2026–2028 分期`, `否`, `表外現金租金（併入路徑；金額未揭露）`],
                          [`承租人自備設備承諾`, `0.5–1.2`, `分期至 2028`, `否`, `併入毛 CapEx`],
                          [`JV 出資承諾（兩案）`, `1.7（已付 0.55）`, `2026 內完成入夥，後續另需增資`, `否`, `JV／策略投資出資`],
                          [`VIE 最大損失暴露（租賃預付）`, `0.108`, `—`, `否`, `不另列`],
                          [`可轉債 2033（9/17 期後）`, `3.7`, `2033-04 到期；2.875%`, `期後`, `股權／可轉債來源＋基本利息`],
                          [`可用信用額度（RCF＋DDTL 未動用）`, `10.01`, `DDTL 4.0 可動用至 2027-06`, `否（來源）`, `不預先計入；為填補缺口的第一順位工具`]
                        ].map(([e, t, n, r, i]) => (0, $.jsxs)(`tr`, {
                          className: `border-b border-border/70`,
                          children: [(0, $.jsx)(`td`, {
                            className: `py-1.5 pr-3 font-medium`,
                            children: e
                          }), (0, $.jsx)(`td`, {
                            className: `py-1.5 pr-3 text-right font-mono tabular-nums`,
                            children: t
                          }), (0, $.jsx)(`td`, {
                            className: `py-1.5 pr-3 text-muted`,
                            children: n
                          }), (0, $.jsx)(`td`, {
                            className: `py-1.5 pr-3 text-muted`,
                            children: r
                          }), (0, $.jsx)(`td`, {
                            className: `py-1.5 text-muted`,
                            children: i
                          })]
                        }, e))
                      })]
                    })
                  }), (0, $.jsxs)(`div`, {
                    className: `mt-3 space-y-2 text-xs leading-relaxed text-muted`,
                    children: [(0, $.jsxs)(`p`, {
                      children: [(0, $.jsx)(`span`, {
                        className: `font-medium text-fg`,
                        children: `未起租 $35.5bn 是契約，不是或有事項。`
                      }), `ASC 842 規定租賃在 commencement 才入表；已簽約、機房未交付 → 表外揭露。10-Q 另把單站 $14.7bn 上限與 355 MW 造價計租排除在 35.5 之外，因「金額具重大不確定性」——三者相加才是承諾全貌，但後兩者金額浮動。`]
                    }), (0, $.jsx)(`p`, {
                      children: `毀約代價 10-Q 未揭露。起租前為開發協議（走掉≈房東已投入造價、開發費）；起租後接近 hell-or-high-water，剩餘租金加速到期。房東端另有可驗證訊號：Galaxy $1.4bn 專案融資（80% LTC、36 個月）、Applied Digital $1.59bn 高收益債（7%）——房東的債務是 CoreWeave 租金的鏡像。`
                    }), (0, $.jsxs)(`p`, {
                      children: [(0, $.jsx)(`span`, {
                        className: `font-medium text-fg`,
                        children: `DDTL 攤還與「營運缺口」的關係：`
                      }), `資產層 DDTL 以客戶合約現金流償還，理論上與 RPO 收現同源；模型把還本列為契約性用途、把新 DDTL 動用列為「未籌」。這使營運缺口看起來比公司敘事嚴峻，但避免了「借新還舊自動成立」的隱含假設。可用額度 $10bn 是第一順位填補工具，未預先扣減缺口。`]
                    })]
                  })]
                }), (0, $.jsxs)(Jj, {
                  className: `border-border bg-surface p-4`,
                  children: [(0, $.jsx)(`h3`, {
                    className: `font-display text-base font-semibold`,
                    children: `法說 vs 10-Q vs 8-K：一致、未量化、含糊`
                  }), (0, $.jsx)(`div`, {
                    className: `mt-2 overflow-x-auto`,
                    children: (0, $.jsxs)(`table`, {
                      className: `w-full min-w-[640px] text-xs`,
                      children: [(0, $.jsx)(`thead`, {
                        children: (0, $.jsxs)(`tr`, {
                          className: `border-b border-border text-left text-muted`,
                          children: [(0, $.jsx)(`th`, {
                            className: `py-2 pr-3 font-medium`,
                            children: `項目`
                          }), (0, $.jsx)(`th`, {
                            className: `py-2 pr-3 font-medium`,
                            children: `法說`
                          }), (0, $.jsx)(`th`, {
                            className: `py-2 pr-3 font-medium`,
                            children: `10-Q／8-K`
                          }), (0, $.jsx)(`th`, {
                            className: `py-2 font-medium`,
                            children: `判定`
                          })]
                        })
                      }), (0, $.jsx)(`tbody`, {
                        children: [
                          [`backlog >50% 已開始交付`, `明確；年底達 2/3`, `RPO 41% 於 24 個月內`, `一致方向；桶內分佈仍是 Derived`],
                          [`CapEx 35–39`, `上修；Q3 11.5–13.5`, `H1 16.139；CIP 11.9 顯示 Q3 前載`, `指引聽法說；CIP 增加證實前載`],
                          [`「資產層融資支應多數 CapEx」`, `明確`, `DDTL 出 13.6／可用 10.0；H1 借款淨 11.5`, `一致；但缺口＝借款需求，不因此消失`],
                          [`加權利率降近 300bps`, `Q2 7.8%（含可轉債 pro forma）`, `DDTL 1.0 15%、2.0 11%、5.0 9%、票據 10%`, `法說數字含 1.75–2.875% 可轉債稀釋效果；邊際借款仍 7–9.75%`],
                          [`可轉債「無稀釋」`, `capped call`, `轉換價 $97.85、cap $199.70；37.8m 股潛在`, `capped call 只覆蓋至 $199.7；股數不計，記於檢查`],
                          [`ATM`, `未提`, `8-K：≤35m 股（約 $2.8bn）`, `上限計入來源開關；實際發行未知`],
                          [`舊世代售罄、A100 至 2029`, `明確`, `無`, `殘值 25%／再出租 75%，不因法說上調`],
                          [`鄉村站點曝險`, `未提`, `10-Q 提及紐約州 50 MW 以上暫停令`, `Bernstein 25%／74% 記於檢查；未量化進模型`],
                          [`FCF`, `合約到期時叢集無槓桿、之後皆 upside`, `H1 CFO 3.663、現金購置 14.117`, `五期 UFCF 為負；「之後 upside」在終值`]
                        ].map(([e, t, n, r]) => (0, $.jsxs)(`tr`, {
                          className: `border-b border-border/70`,
                          children: [(0, $.jsx)(`td`, {
                            className: `py-1.5 pr-3 font-medium`,
                            children: e
                          }), (0, $.jsx)(`td`, {
                            className: `py-1.5 pr-3 text-muted`,
                            children: t
                          }), (0, $.jsx)(`td`, {
                            className: `py-1.5 pr-3 text-muted`,
                            children: n
                          }), (0, $.jsx)(`td`, {
                            className: `py-1.5`,
                            children: r
                          })]
                        }, e))
                      })]
                    })
                  })]
                }), (0, $.jsx)(VM, {
                  tag: `Verified`,
                  text: `Q2 2026 10-Q（2026-06-30，2026-08-12 申報；Non-accelerated filer，Q3 10-Q 期限 11/16）：營收 2.575bn +112%（H1 4.653）；營收成本 0.879；技術與基礎設施 1.507（含 D&A 1.393、營業租賃成本 0.500、變動租賃 0.150）；GAAP 營損 −0.049；利息費用 0.640（H1 1.176；現金利息 0.806、資本化 0.176）；淨損 −0.626（H1 −1.366）；稀釋 EPS −1.14；SBC 0.165；Adj. EBITDA 1.510（59%）；Adj. 營業利益 0.128。`
                }), (0, $.jsx)(VM, {
                  tag: `Verified`,
                  text: `資產負債：現金 5.524、受限 0.873+0.507、有價證券 0.015、可用額度 10.014（總流動性 15.553）；PP&E 46.736（CIP 11.9）；營業 ROU 16.595；遞延收入 9.692（流動 2.686）；追索債務淨 31.405、非追索 3.663；本金合計 35.551（DDTL 1.0 1.30／2.0 3.19／2.1 3.00／3.0 2.22／5.0 1.10／4.0 非追索 2.84；票據 2030 2.0、2031 1.75+2.75、2032 1.25+EUR 2.28、可轉債 2.59+4.00；OEM 4.22+0.88；Magnetar 0.19）。股東權益 5.024。`
                }), (0, $.jsx)(VM, {
                  tag: `Verified`,
                  text: `現金流（H1）：CFO 3.663；現金購置 PP&E 14.117（另 JV 出資 0.550、策略投資 0.138）；借款 16.747、還款 5.219、capped call 0.492、私募股權 2.982（NVIDIA 2.0 @ $87.20、4 月 1.0 @ $109）。非現金：PP&E 相關負債 9.796、OEM 重分類 1.469、新增營業 ROU 8.359。CapEx 口徑（含融資租賃、扣 CIP 變動）Q2 9.352、H1 16.139。`
                }), (0, $.jsx)(VM, {
                  tag: `Verified`,
                  text: `RPO 103.7bn（backlog 104.2）：41% 於 2028-06 前、39% 於 25–48 月、20% 於 49–78 月。客戶集中：A 36%、B 26%、C 10%。OpenAI 承諾最高 6.5bn 至 2031-05；Meta 最高 21.0bn（至 2032-12／2032-04 選擇權）；Jane Street 約 6.0bn（私人公司，另 1.0bn 策略投資）。`
                }), (0, $.jsx)(VM, {
                  tag: `Verified`,
                  text: `租賃與承諾：在帳未折現 29.135（營業）+0.235（融資），2030 後 19.023；加權剩餘租期 12 年、折現率 10%。未起租租賃未折現 35.5（2026–2029 起租，7–16 年）；單站 393 MW 未交付、按造價計租、上限 14.7／16 年；355 MW 按造價計租（金額未定）；承租人自備設備 0.5–1.2；JV 出資承諾 1.7；VIE 暴露 0.108。`
                }), (0, $.jsx)(VM, {
                  tag: `Verified`,
                  text: `9/17 8-K：$3.0bn（+0.5bn 選擇權，實際定價 3.7bn）2033 可轉債，2.875%、轉換價 $97.85（溢價 22.5%）、capped call 上限 $199.70；股權分銷協議最多 35m 股（含 collared forward），30 日內不得發行。9/17 收盤 $79.88。`
                }), (0, $.jsx)(VM, {
                  tag: `Interested-party`,
                  text: `法說（8/11，Intrator／Agrawal）：主動電力 1.5 GW（Q2 +近 500 MW，6 月 +300 MW）；簽約 3.7 GW → 4.2 GW（8/11）；另 1.5 GW powered land／選擇權／LOI；2030 ≥8 GW；海外簽約 >1 GW（印尼 360 MW，約 18 個月後）。近期產能實質售罄；7 月 SKU 漲價約 25%；新約貢獻率高 5–10 pt；A100 續約至 2029。全年營收 12.4–13.2、調整後營益 0.96–1.15、CapEx 35–39、期末 ARR 18.5–19.5；Q3 營收 3.45–3.6、調整後營益 0.20–0.26、利息 0.86–0.94、CapEx 11.5–13.5；Q4 調整後營益率低雙位數。managed inference ARR 年底 ≥0.25bn；其他服務 ARR >0.4bn。DDTL 5.5 為短天期合約融資。`
                }), (0, $.jsx)(VM, {
                  tag: `Interested-party`,
                  text: `9/17 可轉債投資人簡報：Q2 pro forma 總債務 38.6bn；加權利率 FY23 14.9% → Q2 8.3% → pro forma 7.8%（不含 DDTL 5.5）；總債務／backlog 0.4x；歷年 backlog 9.9／15.1／66.8／104.2；「多數 CapEx 由資產層融資支應」；契約貢獻率同時爬坡壓縮、穩定後 mid-20%s。`
                }), (0, $.jsx)(VM, {
                  tag: `Verified`,
                  text: consSourceTxtQ().peers
                }), (0, $.jsx)(VM, {
                  tag: CONSENSUS.priceTarget.tag,
                  text: consSourceTxtQ().targets
                }), (0, $.jsx)(VM, {
                  tag: `Interested-party`,
                  text: `房東揭露（站點租賃檢驗用）：Galaxy Helios 526 MW critical IT、15 年＋2×5 年延長、年均租金 >$1bn（2026-07-06；Phase I 133 MW 於 2026Q2 起租）；Applied Digital Polaris Forge 1 三份租約合計 400 MW、約 $11bn、約 15 年（2025-08-29）；Core Scientific 約 590 MW、$10.2bn、12 年（2025-02-26）。具名合計 1,516 MW、36.2bn，加權每 MW 年租金約 $1.70m。Q4 2025 法說：年底主動電力 >850 MW、2025 CapEx 14.9bn。`
                }), (0, $.jsx)(VM, {
                  tag: `Verified`,
                  text: `10-Q Note 3：兩個未合併 JV 承諾最高 1.7bn，預計 2026 年內履行，入夥後須依協議追加出資；6/30 帳面 0.479。新澤西 JV 持股 35%，租約完工後起租、15 年、租金依建造成本計。Note 8：單站 393 MW 未交付（2026、2028 分期）、按造價計租上限 14.7／16 年；另 355 MW 按部分造價計租、2026–2028 分期交付。`
                }), (0, $.jsx)(VM, {
                  tag: `User-provided`,
                  text: `5Y CDS（2026-09）：中價 720bps、Bid/Ask 690／750、近期區間 680–855。對照 7/28 峰值約 855（當時市場推估五年累積違約機率約 50%）與 6 月低點 452。`
                }), (0, $.jsx)(VM, {
                  tag: `Analogy`,
                  text: `站點 MW：Helios Phase I 133 MW critical IT（Galaxy 7/6 公告）、Phase II 260 MW 施工中、契約全站 800 MW（Galaxy）；Polaris Forge 1 250+150+150 MW（Applied Digital）；Core Scientific 約 590 MW 契約；Kenilworth／Lancaster／印尼規模為公告或推定。10-Q 之 393 MW／355 MW 未交付與上述站點的對應為推定。`
                }), (0, $.jsx)(VM, {
                  tag: `Assumed`,
                  text: `FY27–FY30 毛 CapEx 45／50／50／45（對應 +1.2–1.7 GW/年、約 $30M/MW）；貢獻率 70%；rev/MW 0.0112–0.0115；表外現金租金 0.3→4.6；存量利息路徑；新產能簽約率 100→80%；非算力服務現金 0.1→2.0；違約率 0.5→2.5%；殘值 25%。CDS 855 為 7/28 峰值（6 月低點 452），9 月現值未取得。可比公司倍數未填。RPO 桶內線性分攤為 Derived。`
                }), (0, $.jsxs)(`p`, {
                  className: `text-xs text-muted`,
                  children: [`損益兩平貢獻率 `, hA(d.totals.be * 100), `。終值 `, mA(d.totals.terminal), `bn。本工具是現金資金模型，不是評等預測或投資建議。`]
                })]
              })]
            }), (0, $.jsxs)(Jj, {
              className: `bg-ink text-accent-fg`,
              children: [(0, $.jsx)(`h2`, {
                className: `font-display text-2xl font-semibold`,
                children: `五期融資：瀑布新債 ${mA(d.totals.newDebt)}bn、股權 ${mA(d.totals.equity)}bn（新股 ${Y(d.totals.newShares,2)}bn 股，原股東最終持股約 ${Y(f.shares/(f.shares+d.totals.newShares)*100,0)}%）。連動評價 $${Y(f.call.blended,0)}，結論「${f.call.call}」。`
              }), (0, $.jsxs)(`p`, {
                className: `mt-2 text-sm leading-relaxed text-accent-soft`,
                children: [`營運缺口（不含融資；含 CapEx／租金／利息／JV）為 `, mA(d.totals.operatingGap), `bn，排程還本另 `, mA(d.totals.debtPay), `bn。缺口在需要前一期先融好：依序動用未動用額度、資產層新債（總債務 ≤ `, Y(e.debtBacklog,2), `× backlog）、股權（$`, Y(e.eqPx,2), ` 折價 `, Y(e.eqDisc*100,0), `%）。CDS `, Y(e.cds, 0), ` bps。`]
              })]
        }), (0, $.jsxs)(Jj, {
          className: `p-4`,
          children: [(0, $.jsx)(`div`, {
            className: `text-xs font-medium text-muted`,
            children: `附錄｜資金橋（模型期：自 2026-06-30 起；FY26 全年口徑見「各期收支」頁底）`
          }), (0, $.jsxs)(`p`, {
            className: `mt-2 font-mono text-sm tabular-nums leading-relaxed`,
            children: [`期初 `, Y(y.open), ` + 營運缺口 `, Y(y.operatingGap), ` + 股權／可轉債 `, Y(y.atm), y.facility ? ` + 瀑布融資（新債＋股權）${Y(y.facility)}` : ``, y.debtPay ? ` − 排程還本 ${Y(y.debtPay)}` : ``, ` = 期末 `, Y(y.end), y.ok ? `` : `（對不上）`]
          }), (0, $.jsxs)(`p`, {
            className: `mt-2 text-xs leading-relaxed text-muted`,
            children: [`產能橋：排程 `, Y(b.scheduled), ` − 瓶頸 `, Y(b.bottleneck), ` − 信用損失 `, Y(b.credit), ` = 收現 `, Y(b.collected), `；新簽約現金 `, Y(d.totals.newCash), `（新簽約收入 `, Y(d.totals.newRev), `）。這就是「Backlog 不是現金」。 RPO 滾動：6/30 `, Y(x.begin), ` + Q3 初新增承諾 `, Y(x.q1Net), ` = `, Y(x.q1End), `；五期排程 `, Y(x.convert), `，2030 後仍未認列 `, Y(x.remaining), `。 支出五層：毛 CapEx `, Y(d.totals.gross), `（客戶預付後 `, Y(d.totals.cashCapex), `）＋ 租金 `, Y(d.totals.lease), `（在帳 `, Y(S.onFive), ` + 表外 `, Y(S.cashFive), `）＋ 利息 `, Y(d.totals.interest), `（其中瀑布新債 `, Y(d.totals.overdraft), `）＋ JV／投資 `, Y(d.totals.div), ` ＋ 排程還本 `, Y(d.totals.debtPay), `。表外現金租金五期 `, Y(S.cashFive), ` + 尾 `, Y(S.tail), ` = `, Y(S.mapped), `，對承諾 50.2（35.5＋14.7）差 `, Y(S.gap), `。`]
          })]
        }), (0, $.jsxs)(Jj, {
          className: `border-accent/20 bg-accent/5 p-4`,
          children: [(0, $.jsx)(`div`, {
            className: `text-xs font-medium text-muted`,
            children: SOURCE_ORDER_NOTE
          }), (0, $.jsx)(`div`, {
            className: `mt-3 grid grid-cols-2 gap-2 sm:grid-cols-3 lg:grid-cols-6`,
            children: [
              [`Q2 營收／H1`, `${Y(LATEST_Q.revenue,3)}bn +${hA(LATEST_Q.yoy*100,0)} · H1 ${Y(LATEST_Q.h1Revenue,3)}`],
              [`RPO 桶`, `${Y(LATEST_Q.rpo,1)}bn · 41／39／20（24m／25–48／49–78）`],
              [`CapEx`, `Q2 ${Y(LATEST_Q.capexQ2,1)}、H1 ${Y(LATEST_Q.capexH1,1)} · 全年 35–39`],
              [`債務本金`, `${Y(LATEST_Q.debtPrincipal,1)}bn · 2H26–FY30 攤還 ${Y(DEBT_AMORT.reduce((e,t)=>e+t,0),1)}`],
              [`客戶預付`, `遞延 ${Y(LATEST_Q.deferredTotal,1)}bn · H1 淨流入 ${Y(LATEST_Q.deferredIn,2)}`],
              [`租賃`, `在帳未折現 ${Y(LATEST_Q.onBalanceUndiscounted,1)} · 未起租 ${Y(LATEST_Q.offBalanceLease,1)}＋單站上限 ${Y(LATEST_Q.singleSiteCap,1)}`]
            ].map(([e, t]) => (0, $.jsxs)(`div`, {
              className: `rounded-md bg-card/80 px-3 py-2`,
              children: [(0, $.jsx)(`div`, {
                className: `text-xs text-muted`,
                children: e
              }), (0, $.jsx)(`div`, {
                className: `mt-0.5 font-mono text-xs tabular-nums leading-snug`,
                children: t
              })]
            }, e))
          }), (0, $.jsx)(`div`, {
            className: `mt-2 grid grid-cols-2 gap-2 sm:grid-cols-4`,
            children: [
              [`主動電力`, `1.5 GW（Q2 +${CALL_FACTS.activeAddQ2} MW）· YE >${CALL_FACTS.yeActiveGw} GW · 觀察，不年化`],
              [`簽約電力`, `${CALL_FACTS.contractedGw} GW（8/11）· 另 1.5 GW 土地／選擇權／LOI`],
              [`Q3 指引`, `營收 3.45–3.6 · 利息 0.86–0.94 · CapEx 11.5–13.5`],
              [`9/16–17`, `9/16：Q3 以更高價簽短天期約 · 9/17 可轉債 $3.7bn 2.875% @ $97.85（淨現金約 3.2）、ATM ≤35m 股（→股數 0.5515+0.035=0.5865）、9/17 收盤 $${CALL_FACTS.price0917}；最新 ${CALL_FACTS.priceDate} 收盤 $${CALL_FACTS.priceLast}`]
            ].map(([e, t]) => (0, $.jsxs)(`div`, {
              className: `rounded-md bg-card/80 px-3 py-2`,
              children: [(0, $.jsx)(`div`, {
                className: `text-xs text-muted`,
                children: e
              }), (0, $.jsx)(`div`, {
                className: `mt-0.5 font-mono text-xs tabular-nums leading-snug`,
                children: t
              })]
            }, e))
          })]
            })]
          })]
        })]
      })]
    })]
  })
}
