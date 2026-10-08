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
    cvCap: SCENARIOS[e].cvCap,
    delayMonths: SCENARIOS[e].delay, // v0.2
    billableOpen: SCENARIOS[e].bOpen,
    m: {
      ...t.m,
      accepted: [...SCENARIOS[e].acc],
      billable: [...SCENARIOS[e].bil],
      revMW: [...SCENARIOS[e].rev]
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
          }), `，加上新融資本身的利息 `, mA(d.years.reduce((e, t) => e + t.newDebtInt, 0)), `bn，由新債 `, mA(d.totals.newDebt), `bn、可轉債 `, mA(d.totals.convNew), `bn、股權 `, mA(d.totals.equity), `bn（新股 `, Y(d.totals.newShares, 2), `bn 股）`, d.totals.junk > .05 ? `、高息債 ${mA(d.totals.junk)}bn` : ``, ` 補足。每 MW 年 EBITDA 約 `, (0, $.jsxs)(`span`, {
            className: `font-medium text-accent-fg`,
            children: [`$`, Y(d.m.revMW[4] * (e.revScale ?? 1) * 1e3 * d.m.util[4] / 100 * d.years[4].ebM, 1), `m`]
          }), `，回收一個 MW 的 GPU（$`, Y(e.a.costMW[4] * (e.capexScale ?? 1), 0), `m、`, e.gpuLife, ` 年、WACC `, hA(o.wacc * 100, 0), `）每年需要 `, (0, $.jsxs)(`span`, {
            className: `font-medium text-accent-fg`,
            children: [`$`, Y(e.a.costMW[4] * (e.capexScale ?? 1) * o.wacc / (1 - Math.pow(1 + o.wacc, -e.gpuLife)), 1), `m`]
          }), `。目標價 `, (0, $.jsxs)(`span`, {
            className: `font-medium text-accent-fg`,
            children: [`$`, Y(f.call.blended, 1)]
          }), `（情境區間 $${Y(TR.A[0], 1)}–$${Y(TR.A[1], 1)}），較現價 `, f.call.upside >= 0 ? `+` : ``, hA(f.call.upside * 100, 0), `，結論「`, f.call.call, `」。`, (0, $.jsx)(tipQ, {
            t: `三個判斷依據：(1) 預付款能否讓 backlog 變成現金——營收由已連網 MW × 每 MW 年收入（Tokenomics 正向推導，不用公司 ACV）驅動，客戶預付在建置前先收現、之後認列為非現金營收；(2) 單位經濟——每 MW 年 EBITDA（每 MW 年收入×利用率×EBITDA 率）是否高於 GPU 的年化資本回收（每 MW 成本×資本回收係數）；(3) 融資——缺口依序由預付、現金、資產擔保融資、可轉債、ATM 股權補足，最後才是高息債；市場願意以多少價格吸收多少新股，是估值最敏感的假設之一（見目標價頁敏感度表）。${PERIODS[0]} 欄為年初至今實際＋模型期。`,
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
                tip: `三個情境依對照表 r1 第 4 節第 1 條：保守＝只交付已簽約電力（>3.5 GW，÷1.2 換 MW-IT）；基準＝2026 年底 5 GW 合約目標依歷史併網速度實現；積極＝2027 起每年 >1 GW 部署全數實現。每 MW 年收入隨情境取 Tokenomics 正向推導三情境值（保守／基準／積極）。毛 CapEx 由 MW 公式自動增減；FY31 預建隨情境。切換情境只會改寫已連網／可計費 MW、每 MW 年收入、FY31 預建與表外租金；你手動調整的其他數字會保留，按頁首「重設」才回到預設值。`,
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
                children: [`目前：`, e.scenario === `custom` ? `自訂` : SCENARIOS[e.scenario]?.label, ` · ${PERIODS[4]} 年底已連網 `, Y(d.m.accepted[4], 0), ` MW · 模型期 CapEx `, Y(d.years.reduce((e, t) => e + t.gross, 0), 0), `bn`]
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
                  hint: `${Y(LATEST_Q.rpo, 1)} 季報 RPO（只作對照，不驅動營收）`,
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
                  label: `${CALQ.valuationMD} Billable MW`,
                  hint: SC_BRM === `converge` ? `以 ${CALQ.filedQLabel} 實際營收年化 ÷ 每 MW 年收入校準（隨情境）` : `Assumed`,
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
                  hint: `關掉則用期末存量全期化。打開用（評價日 ${e.billableOpen} + 當期期末）/2，預設打開。`
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
                  hint: `${PERIODS[0]} Accepted/Billable 不低於具名站點加總；不足部分進殘差列。`
                })]
                }),
                (0, $.jsx)(accQ, {
                  title: `C｜利潤率（EBITDA 單一來源）`,
                  sum: `EBITDA ${hA(e.ebStart*100,0)}→${hA(e.ebSteady*100,0)}`,
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
                  label: `穩態 EBITDA 率（${PERIODS[4]}）`,
                  hint: `${hA(e.ebSteady*100,1)} · 線性變動；敏感度取可觀察 neocloud 區間（IREN 約 35%、CRWV 約 59%）`,
                  children: [(0, $.jsx)(IM, {
                    value: e.ebSteady * 100,
                    onChange: e => w({
                      ebSteady: e / 100
                    }),
                    step: .5
                  }), (0, $.jsxs)(`p`, {
                    className: `text-xs text-muted`,
                    children: [`損益與資金共用；資金端用 EBITDAR 率＝EBITDA 率＋租金÷營收`, (0, $.jsx)(tipQ, {
                      t: `EBITDA 率路徑不取自每 MW 推導路徑 A 的加成（其隱含約 75%，高於可觀察水準），改以可觀察 neocloud 區間為準：IREN FY26 調整後 EBITDA 率約 35%、CRWV Q2 2026 約 59%（v0.1a 總帳）。起始取 ${TXQ.ebStartSource}，穩態取區間中點 47%。EBITDA 已扣營業租賃成本；資金模型把租金列在支出端，所以 EBITDAR 率須把租金加回，兩邊才不會重複扣除。`,
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
                  label: `模型期後一年新增 MW（${PERIODS[4]} 預建用）`,
                  hint: `模型期後一年新增 MW（${PERIODS[4]} 預建用）；隨情境`,
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
                  hint: `季報現金 ${Y(LATEST_Q.cash, 3)}（受限現金不計）`,
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
                  label: `計入期後股權／可轉債淨現金 $${Y(e.atm, 2)}bn`,
                  hint: `${CALL_FACTS.postQShortDated || ``}。淨現金 ${e.atm}bn 記於首期模型部分，不算進「營運缺口」；新可轉債同時是負債（見既有債務）。`
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
                  hint: `未動用額度 ${e.facility}bn（資產擔保定期貸款，2026-07 簽約）。瀑布第一順位，不受 債務／backlog 上限限制（已承諾額度）。關閉＝不動用額度，直接進入新債與股權。`
                }),
(0, $.jsx)(RM, {
                  checked: e.includeDebt,
                  onChange: e => w({
                    includeDebt: e
                  }),
                  label: `債務排程攤還（季報到期表）`,
                  hint: `可轉債依到期日以累積本金還本，屬契約性支出，預設打開。關閉＝假設全額再融資。不含新借款。`
                }),
(0, $.jsx)(LM, {
                  label: `債務／backlog 上限`,
                  hint: `資產擔保融資容量：總債務 ≤ 此倍數 × backlog（預設 ${e.debtBacklog}x，[Assumed]；評價日實際約 0.27x）`,
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
                  hint: `新債利率 = 基準 + max(0, CDS−${e.cdsBaseBp})×${e.cdsPassThrough}。沒有 CDS 報價時不適用；預設關閉。不進入客戶違約。`
                }),
(0, $.jsxs)(`div`, {
                className: `mt-5 rounded-lg border border-border bg-surface p-3`,
                children: [(0, $.jsx)(`div`, {
                  className: `text-sm font-medium`,
                  children: `${COMPANY_DATA.meta.company} 5Y CDS`
                }), (0, $.jsxs)(`p`, {
                  className: `mt-1 text-xs leading-relaxed text-muted`,
                  children: e.cds == null ? [`不適用：`, e.cdsDate, `。`] : [`中價 `, Y(e.cds, 0), ` bps（`, e.cdsDate, `）；Bid/Ask `, Y(e.cdsBid, 0), `／`, Y(e.cdsAsk, 0), `；近期區間 `, Y(e.cdsLo, 0), `–`, Y(e.cdsHi, 0), `。這是公司自身信用，不是客戶違約率。`]
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
                    t: `左圖：各期營運來源、融資來源與用途。右圖：紅色面積為「融資前累積現金」＝若不做任何新融資，評價日現金加上各期營運來源（含客戶預付）與期後股權／可轉債、減去所有支出後的累積結果，負值即需要外部資金的規模；線條為瀑布累計補入的新債、股權與高息債。期前融資後的期末現金固定在最低現金 2.0bn，因此不再畫出。`,
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
                        return d.years.map(e => (t += e.newDebt + e.convNew, n += e.equity, r += e.junk, {
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
                  tip: `支出五層：① 毛 CapEx（由已連網 MW 推導；客戶預付在來源端列示）→ ② 租金（在帳＝季報到期表固定值；表外＝已簽約未起租 ${Y(LATEST_Q.offBalanceLease, 1)}bn 的現金路徑）→ ③ 利息（存量債務依既有債務推算；瀑布新債另計利息）→ ④ 策略投資出資 → ⑤ 排程還本（季報到期表）。藍色可編輯格為輸入；灰底列為計算結果。本頁輸入為首期模型部分的金額，${PERIODS[0]} 欄在「各期收支」頁才加回年初至今實際。`
                }), (0, $.jsx)(BM, {
                  rows: [
                    [`支出（可編輯輸入；${PERIODS[0]} 欄為${CALQ.stubWord}模型值）`, null],
                    [`　每 MW 建置成本`, e.a.costMW, (e, t) => E(`costMW`, e, t), 1, void 0, `${TXQ.costMwNote}。全年指引不用來回推。`, `US$m/MW`],
                    [`　本期新增 MW`, d.years.map(e => e.mwNew), void 0, void 0, `calc`, `＝期末 Accepted − 期初。${PERIODS[0]} 期初為上一財年底 ${e.mwYearEnd[PERIOD_FY[0] - 1]} MW。`, `MW`],
                    [`　次期新增 MW`, d.years.map(e => e.mwNext), void 0, void 0, `calc`, `下一期新增量；${PERIODS[4]} 欄取左欄「模型期後一年新增 MW」。`, `MW`],
                    [`　全年毛 CapEx（公式）`, d.years.map(e => e.capexFull), void 0, void 0, `calc`, `＝(本期新增×(1−λ)＋次期新增×λ)×每 MW 成本。${PERIODS[0]} 對照公司全年指引 ${CALL_FACTS.capexLo}–${CALL_FACTS.capexHi}（${TXQ.capexGuideSource}）。`],
                    [`　成長型 CapEx（模型期）`, d.years.map(e => e.capexGrowth), void 0, void 0, `calc`, `${PERIODS[0]} 欄＝MAX(全年公式, 下限 ${e.capexFloorFY0}) − 年初至今 ${Y(ACTUAL_1H.capex, 3)}。下限代表當年已下單、無論情境都會發生的支出。`],
                    [`　GPU 汰換 CapEx`, d.years.map(e => e.refresh), void 0, void 0, `calc`, `${e.refreshSteady ? `已連網 MW 不再增加的期間（觸頂後）及 ${PERIODS[4]}：穩態汰換＝期間平均已連網 MW × 每 MW GPU 成本 ÷ 壽命 ${e.gpuLife} 年（建築與電力屬租賃不計）；其餘期間：` : ``}＝(本年 − 經濟壽命) 那一年新增的 MW × 每 MW 成本。壽命 ${e.gpuLife} 年；批次來源見左欄年底主動電力（${Object.entries(e.mwYearEnd).map(([y, m]) => `FY${String(y).slice(2)} 年底 ${m} MW`).join(`、`)}），汰換落在模型期之後者不出現。三情境相同。`],
                    [`① 毛 CapEx（模型期）`, d.years.map(e => e.gross), void 0, void 0, `tot`, `＝成長型＋汰換。含 OEM 融資的非現金部分。`],
                    [`　閒置資本（已支出未產生收入，期末）`, d.years.map(e => e.idleCap), void 0, void 0, `calc`, `GPU 成長型資本支出照原併網時程（已採購、交貨等電），計費延後 ${multTxt(e.delayMonths ?? 0)} 個月：＝原時程累計 − 已投入使用累計；折舊自投入使用時點起算（v0.2）。`],
                    [`　對照：v1.4 手動值`, d.years.map(e => e.capexOld), void 0, void 0, void 0, `CRWV 模板舊版手動值（Oracle 不適用，0）。`],
                    [`　客戶預付率`, e.a.customerFund.map(e => e * 100), (e, t) => E(`customerFund`, e, t / 100), void 0, void 0, `＝有預付的合約比例 × 預付占資本支出比（${TXQ.prepayCoverNote}）。只降當期外部融資需求、形成合約負債，不降專案總成本。`, `%`],
                    [`② 表外現金租金（未起租）`, e.a.newLease, (e, t) => E(`newLease`, e, t), void 0, void 0, `對應季報已簽約未起租租賃 ${Y(LATEST_Q.offBalanceLease, 1)}bn（${TXQ.offBalanceLeaseTerm}）的現金支付路徑 [Derived]。`],
                    [`　表外租金（延誤連動後）`, d.years.map(e => e.offLease), void 0, void 0, `calc`, `上列原排程中 ${pctQ(e.delayLink ?? 0)} 的起租隨建設延誤 ${multTxt(e.delayMonths ?? 0)} 個月後移（開發商交付晚），其餘照原時程；租金合計用此列（v0.2）。`],
                    [`② 在帳現金租金（季報固定）`, [...LEASE_CASH_ON_BAL], void 0, void 0, void 0, `季報到期表：${LEASE_CASH_ON_BAL.map(x => Y(x, 2)).join(`／`)}，之後尚有 ${Y(LEASE_AFTER_FY30, 2)}。`],
                    [`③ 存量債務利息（既有債務推算）`, d.years.map(e => e.intStock), void 0, void 0, `calc`, `＝平均本金（依到期表遞減）× 加權有效利率 ${hA(DBT_R * 100, 1)} × 期間長度 ＋ 期後新發可轉債利息 ＋ 首期校準 ${e.intCal}。明細見「既有債務」分頁。`],
                    [`　對照：v1.4 手動值`, d.years.map(e => e.intOld), void 0, void 0, void 0, `CRWV 模板舊版手動值（Oracle 不適用，0）。`],
                    [`③ 新債利息（瀑布，計算）`, d.years.map(e => e.newDebtInt), void 0, void 0, `calc`, `＝新債利率 × 期間長度 × (期初新債餘額 ＋ 本期舉借)。期前融資：本期舉借在期初到位，當期全額計息。`],
                    [`④ JV 已承諾餘額`, d.years.map(e => e.jvC), void 0, void 0, void 0, `季報未揭露 JV 出資承諾（不適用）。`],
                    [`④ JV 後續增資＋策略投資`, e.a.div, (e, t) => E(`div`, e, t), void 0, void 0, `收購與策略投資，未揭露計畫 [Assumed]。年初至今 ${Y(ACTUAL_1H.jv, 3)}。`],
                    [`⑤ 排程還本（季報到期表）`, e.includeDebt ? [...DEBT_AMORT] : DEBT_AMORT.map(() => 0), void 0, void 0, void 0, `可轉債以到期累積本金還本；關閉開關＝假設全額再融資。`],
                    [`來源`, null],
                    [`非算力服務營收`, e.services, (n, r) => w({
                      services: e.services.map((e, t) => t === n ? r : e)
                    }), void 0, void 0, `${TXQ.otherRevNote}。損益營收與資金現金共用這一列。`],
                    [`EBITDA 率（路徑）`, d.years.map(e => e.ebM * 100), void 0, void 0, `calc`, e.ebitdaBasis === `ebitdar` ? `＝EBITDAR 率 − 租金÷營收（租金為固定成本）。` : `由左欄起始與穩態值線性推得。`, `%`],
                    [`EBITDAR 率（計算）`, d.years.map(e => e.cashMargin * 100), void 0, void 0, `calc`, e.ebitdaBasis === `ebitdar` ? `三情境共用：起始 ${(e.ebStart*100).toFixed(1)}%＋基準租金比 ${(e.ebitdarAdj[0]*100).toFixed(1)}% → 穩態 ${(e.ebSteady*100).toFixed(1)}%＋${(e.ebitdarAdj[1]*100).toFixed(1)}%，線性；使基準情境起點與穩態 EBITDA 率維持輸入值，低營收情境承擔固定租金。` : `＝EBITDA 率＋租金÷營收，即租金前的 EBITDA 率（業界稱 EBITDAR）。租金在支出端另列，故須加回。`, `%`],
                    [`非算力服務現金（計算）`, d.years.map(e => e.legacy), void 0, void 0, `calc`, `＝服務營收 × EBITDAR 率。`],
                    [`新簽約現金（計算）`, d.years.map(e => e.newCash), void 0, void 0, `calc`, `＝未被期初 RPO 占用的產能 × 新產能簽約率 × (1−信用損失率) × EBITDAR 率。`]
                  ]
                })]
              }), n === 1 && (0, $.jsxs)(`div`, {
                className: `space-y-3`,
                children: [(0, $.jsx)(hdrQ, {
                  title: `各期收支（類現金流量）`,
                  tip: `${PERIODS[0]} 欄＝年初至今實際（季報）＋模型期，所以可直接對照公司全年指引：營收 ${REV_GUIDE_TXT}、CapEx ${CALL_FACTS.capexLo}–${CALL_FACTS.capexHi}。${PERIODS[1]} 以後為純模型。負數以括號表示。現金橋：期初 ＋ 營運缺口 ＋ 股權／可轉債 ＋ 未動用額度 − 排程還本 ＝ 期末。`
                }), (0, $.jsx)(BM, {
                  rows: [
                    [`收入（產能約束）`, null],
                    [`排程 RPO（對照）`, d.years.map(e => e.scheduled), void 0, void 0, void 0, `＝(評價日 RPO ${Y(e.rpoOpen, 1)} ＋ 期後新增 ${Y(e.rpoPendingAdd, 1)}) × 本期桶權重 ÷ 權重合計 × 五期認列比例。只作對照與產能瓶頸旗標，不驅動營收。`],
                    [`容量上限`, d.years.map(e => e.capacity), void 0, void 0, void 0, `＝平均在役 MW × 每 MW 年收入 × 利用率 × 期間係數。只看機房，不看合約。`],
                    [`平均在役 MW`, d.years.map(e => e.avgBillable), void 0, void 0, void 0, `＝(期初 Billable ＋ 期末 Billable) ÷ 2。可在左欄關閉，改用期末存量全期化。`, `MW`],
                    [`產能瓶頸`, d.years.map(e => e.bottleneck), void 0, void 0, `calc`, `＝MAX(0, 排程 − 容量)。合約有、機房沒有，本期收不到且不遞延到下期（保守處理）。`],
                    [`期初 RPO 轉換收入`, d.years.map(e => e.revenue), void 0, void 0, void 0, `＝排程 − 瓶頸。瓶頸>0 時恆等於容量上限。`],
                    [`新簽約收入（計算）`, d.years.map(e => e.newRev), void 0, void 0, `calc`, `＝MAX(0, 容量 − 排程) × 新產能簽約率。FY28 起模型收入主要來自這裡——尚未簽署的合約。`],
                    [`未售產能`, d.years.map(e => e.unsold), void 0, void 0, void 0, `蓋好但賣不掉。MW 驅動時新產能簽約率 100%，此列為 0。`],
                    [`信用損失`, d.years.map(e => e.loss), void 0, void 0, void 0, `＝收入 × 違約率 × (1−回收率)。前三大客戶占 72% 營收。`],
                    [`損益用算力收入`, d.years.map(e => e.isRev), void 0, void 0, `tot`, `＝RPO 轉換 ＋ 新簽約。損益與評價頁用的是同一個數字。`],
                    [`來源（FY26 欄＝1H 實際現金流＋下半年模型）`, null],
                    [`Ⓐ0 年初至今實際營運現金流（CFO）`, d.years.map((e, t) => t === 0 ? e.fyCfo : 0), void 0, void 0, void 0, `季報實際值 ${Y(ACTUAL_1H.cfo, 3)}，已含年初至今的利息、租金與客戶預付（遞延營收增加 ${Y(LATEST_Q.deferredIn, 3)}），因此下方②③Ⓓ的 ${PERIODS[0]} 欄只含模型期。`],
                    [`　EBITDA 率（損益、資金共用）`, d.years.map(e => e.ebM * 100), void 0, void 0, void 0, e.ebitdaBasis === `ebitdar` ? `＝EBITDAR 率 − 租金÷營收。` : `起始 → 穩態線性爬升。`, `%`],
                    [`　EBITDAR 率（EBITDA 率＋租金÷營收）`, d.years.map(e => e.cashMargin * 100), void 0, void 0, void 0, `租金前的 EBITDA 率。租金在支出②另列，所以這裡加回，避免重複扣除。`, `%`],
                    [`Ⓐ RPO 現金（下半年起）`, d.years.map(e => e.rpoCash), void 0, void 0, void 0, `＝收現 × EBITDAR 率。`],
                    [`Ⓑ 新簽約現金`, d.years.map(e => e.newCash)],
                    [`Ⓒ 非算力服務現金`, d.years.map(e => e.legacy)],
                    [`Ⓒ4 減：延誤罰則（營業費用）`, d.years.map(e => -e.delayPen), void 0, void 0, void 0, `建設延誤期間應計費而未計費營收 × 罰則比例 ${pctQ(e.delayPenalty ?? 0)}（公司未揭露合約條款，預設 0；v0.2）。`],
                    [`Ⓓ 客戶預付（${CALQ.stubWord}起）`, d.years.map(e => e.external), void 0, void 0, void 0, `年初至今的遞延營收淨增 ${Y(LATEST_Q.deferredIn, 3)} 已含在 CFO 內，不重複計入。`],
                    [`Ⓓ2 減：預付認列（非現金營收）`, d.years.map(e => -e.prepayRecog), void 0, void 0, void 0, `＝期初合約負債 ÷ 認列年數 ${e.prepay.recogYears} × 期間長度；這部分營收已在預付時收現，不重複計入。`],
                    [`　合約負債期末（客戶預付餘額）`, d.years.map(e => e.clEnd), void 0, void 0, void 0, `期初 ${Y(e.prepay.openBalance, 3)}（${TXQ.prepayOpenNote}）＋預付流入 − 認列。`],
                    [`營運來源合計`, d.years.map((e, t) => t === 0 ? e.fySourcesOp : e.sourcesOp), void 0, void 0, `tot`],
                    [`Ⓔ 股權／可轉債（融資）`, d.years.map((e, t) => t === 0 ? e.fyEquity : e.atm), void 0, void 0, void 0, `${PERIODS[0]}＝年初至今股權 ${Y(ACTUAL_1H.equity, 3)}（${TXQ.ytdEquityNote}）＋期後 ${Y(e.atm, 2)}。`],
                    [`Ⓕ 年初至今實際借款（融資）`, d.years.map((e, t) => t === 0 ? e.fyBorrow : 0), void 0, void 0, void 0, `季報：年初至今借款 ${Y(ACTUAL_1H.borrow, 3)}。`],
                    [`Ⓖ 瀑布：新債（額度＋資產層）`, d.years.map(e => e.newDebt), void 0, void 0, void 0, `先用未動用額度，再用資產層新債；總債務不得超過 債務／backlog 上限。`],
                    [`Ⓖ2 瀑布：可轉債`, d.years.map(e => e.convNew), void 0, void 0, void 0, `資產擔保融資用罄後、股權之前；每年上限 ${Y(e.cvCap ?? 0, 1)}bn（隨情境）、票息 ${hA(e.convIssue.coupon * 100, 1)}。`],
                    [`Ⓗ 瀑布：股權募資`, d.years.map(e => e.equity), void 0, void 0, void 0, `債務與可轉債用罄後的殘差，按現價折價發行；每年不超過股權吸收上限。`],
                    [`Ⓘ 瀑布：高息債（股權上限溢出）`, d.years.map(e => e.junk), void 0, void 0, void 0, `股權超過每年吸收上限的部分，以高息債補足；不受 backlog 上限約束。`],
                    [`總來源（含融資）`, d.years.map((e, t) => t === 0 ? e.fySrcTotal : e.sources), void 0, void 0, `tot`],
                    [`用途（現金口徑）`, null],
                    [`① CapEx（1H 現金購置／下半年起毛額）`, d.years.map((e, t) => t === 0 ? e.fyCapexUse : e.gross), void 0, void 0, void 0, `年初至今採現金流量表的現金購置 ${Y(ACTUAL_1H.cashCapex, 3)}；模型期採毛額，客戶預付在來源端列示。`],
                    [`② 租金（在帳＋表外）`, d.years.map(e => e.lease), void 0, void 0, void 0, `${PERIODS[0]} 欄只含模型期；年初至今租金已含在 CFO 內。`],
                    [`③ 利息（含瀑布新債）`, d.years.map(e => e.interest), void 0, void 0, void 0, `${PERIODS[0]} 欄只含模型期；年初至今利息已含在 CFO 內。`],
                    [`④ JV／策略投資`, d.years.map((e, t) => t === 0 ? e.fyDiv : e.div), void 0, void 0, void 0, `${PERIODS[0]}＝年初至今實際 ${Y(ACTUAL_1H.jv, 3)} ＋ 模型期 ${Y(d.years[0].div, 2)}。`],
                    [`⑤ 排程還本`, d.years.map((e, t) => t === 0 ? e.fyDebtPay : e.debtPay), void 0, void 0, void 0, `${PERIODS[0]}＝年初至今實際還款 ${Y(ACTUAL_1H.debtRepaid, 3)} ＋ 模型期到期表 ${Y(DEBT_AMORT[0], 3)}。`],
                    [`⑥ capped call（年初至今實際）`, d.years.map((e, t) => t === 0 ? e.fyCapped : 0)],
                    [`用途合計`, d.years.map((e, t) => t === 0 ? e.fyUsesCash : e.uses), void 0, void 0, `tot`],
                    [`期前融資瀑布`, null],
                    [`融資前現金（扣既有新債利息）`, d.years.map(e => e.preCash), void 0, void 0, void 0, `＝期初現金 ＋ 營運來源 ＋ 可轉債／ATM − 用途（不含本期舉借利息）。`],
                    [`融資需求（補足至最低現金）`, d.years.map(e => e.need), void 0, void 0, `calc`, `＝MAX(0, 最低現金 − 融資前現金)。`],
                    [`期末 backlog`, d.years.map(e => e.backlogEnd), void 0, void 0, void 0, `＝期初 − 排程認列 − 新簽約認列 ＋ 新簽約增量 × 合約年期。期初 ${Y(e.rpoOpen + e.rpoPendingAdd, 1)}（評價日 RPO）。只用於資產擔保融資容量（總債務 ≤ 倍數 × backlog）。`],
                    [`債務上限（債務／backlog）`, d.years.map(e => e.debtCap)],
                    [`期末總債務（既有＋可轉債＋新債）`, d.years.map(e => e.totalDebtEnd), void 0, void 0, `tot`],
                    [`每年股權吸收上限`, d.years.map(e => Number.isFinite(e.eqCap) ? e.eqCap : NaN), void 0, void 0, void 0, `＝現市值 × 上限 % × 期間長度。`],
                    [`新發行股數`, d.years.map(e => e.newShares), void 0, void 0, void 0, `＝股權募資 ÷ 發行價。`, `bn 股`, 3],
                    [`可轉債（瀑布）餘額`, d.years.map(e => e.convEnd), void 0, void 0, void 0],
                    [`高息債餘額`, d.years.map(e => e.junkEnd), void 0, void 0, void 0],
                    [`融資前累積現金`, d.years.map(e => e.preFinCum), void 0, void 0, `tot`, `若不做任何新融資的累積現金；負值＝外部資金需求。`],
                    [`累計新股`, d.years.map(e => e.cumNewShares), void 0, void 0, `tot`, void 0, `bn 股`, 3],
                    [`缺口與現金`, null],
                    [`營運缺口（不含融資與還本）`, d.years.map((e, t) => t === 0 ? e.fyOperatingGap : e.operatingGap), void 0, void 0, `tot`, `本業能不能自給。${PERIODS[0]} 的年初至今部分＝CFO ${Y(ACTUAL_1H.cfo, 3)} − 現金 CapEx ${Y(ACTUAL_1H.cashCapex, 3)} − 策略投資 ${Y(ACTUAL_1H.jv, 3)}（皆為季報實際值）。`],
                    [`期初累積現金`, d.years.map((e, t) => t === 0 ? ACTUAL_1H.cash1231 : d.years[t - 1].cum), void 0, void 0, void 0, `${PERIODS[0]} 自上一年底現金 ${Y(ACTUAL_1H.cash1231, 3)} 起算。`],
                    [`1H 其他／受限現金調節`, d.years.map((e, t) => t === 0 ? e.hPlug : 0), void 0, void 0, void 0, `使年初至今實際流量接回評價日現金餘額；差額來自受限現金變動、匯率與未逐項列出的項目。`],
                    [`期末累積現金`, d.years.map(e => e.cum), void 0, void 0, `tot`, `負值＝尚需向資產擔保融資／可轉債／股權市場籌措的金額。未動用額度 ${Y(e.facility, 3)}bn 未預先扣減。`],
                    [`FY26 全年備忘（認列口徑，對照公司指引）`, null],
                    [`CapEx 認列（年初至今 ${Y(ACTUAL_1H.capex, 3)}＋模型期）`, d.years.map((e, t) => t === 0 ? e.fyGross : e.gross), void 0, void 0, void 0, `公司全年指引 ${CALL_FACTS.capexLo}–${CALL_FACTS.capexHi}（法說會轉述）。`],
                    [`利息（1H ${Y(ACTUAL_1H.interest, 3)}＋下半年）`, d.years.map((e, t) => t === 0 ? e.fyInterest : e.interest), void 0, void 0, void 0, `公司未提供季度利息指引。`],
                    [`租金現金（1H ${Y(ACTUAL_1H.leasePaid, 3)}＋下半年）`, d.years.map((e, t) => t === 0 ? e.fyLease : e.lease)]
                  ]
                }), (0, $.jsxs)(`p`, {
                  className: `text-xs leading-relaxed text-muted`,
                  children: [`${PERIODS[0]} 現金橋：上年底現金 `, Y(ACTUAL_1H.cash1231), ` ＋ 年初至今營運 `, Y(d.totals.h1Op), ` ＋ 年初至今融資 `, Y(d.totals.h1Fin), ` − 年初至今還款 `, Y(ACTUAL_1H.debtRepaid), ` ＋ 其他／受限現金 `, Y(d.totals.h1Plug), ` ＝ 評價日現金 `, Y(e.cash), `；再加模型期缺口 `, Y(d.years[0].gap), ` ＝ ${PERIODS[0]} 期末 `, Y(d.years[0].cum), `。`]
                })]
              }), n === 2 && (0, $.jsxs)(`div`, {
                className: `space-y-3`,
                children: [(0, $.jsx)(hdrQ, {
                  title: `收入／產能輸入`,
                  tip: `營收 ＝ 平均在役 MW × 每 MW 年收入 × 利用率 × 期間長度（MW 驅動，新產能簽約率 100%）。每 MW 年收入取 Tokenomics 正向推導三情境（11.62／17.40／24.20 US$m/MW-IT），不用公司 ACV（$20–25M 只作對照）。路徑 B 已含可計費利用率，所以利用率預設 100%。RPO 排程只作對照與產能瓶頸旗標。${PERIODS[0]} 欄的 MW 為年底存量、收入為模型期金額。`
                }), (0, $.jsx)(BM, {
                  rows: [
                    [`Accepted MW（期末主動電力）`, d.m.accepted, (e, t) => D(`accepted`, e, t), void 0, void 0, `YE26 指引 >1,850 MW；2030 目標 ≥8,000 MW。引擎會強制單調不減。`, `MW`],
                    [`Billable MW`, d.m.billable, (e, t) => D(`billable`, e, t), void 0, void 0, `引擎會強制不超過 Accepted。`, `MW`],
                    [`Billable MW（延誤後，計費用）`, d.years.map(e => e.billDelayed), void 0, void 0, `calc`, `＝上列往後平移建設延誤 ${multTxt(e.delayMonths ?? 0)} 個月（以期間長度線性內插；評價日之前取期初校準值）；營收依此列（v0.2）。`, `MW`],
                    [`利用率`, e.m.util, (e, t) => D(`util`, e, t), void 0, void 0, `法說稱「近期產能實質售罄」，本模型不擬合為 100%。`, `%`],
                    [`每 MW 年收入`, e.m.revMW, (e, t) => D(`revMW`, e, t), 5e-4, void 0, `Tokenomics 正向推導（data/permw_tokenomics_20261007.json）；隨情境：保守 0.01162／基準 0.0174／積極 0.0242 [Derived]。`, `US$bn/MW`],
                    [`新產能簽約率`, e.m.fill, (e, t) => D(`fill`, e, t), 1, void 0, `把這欄調成 0，就能看到只靠期初 RPO 的缺口有多大——最重要的壓力測試。`, `%`],

                  ]
                })]
              }), n === 3 && (0, $.jsxs)(`div`, {
                className: `space-y-3`,
                children: [(0, $.jsx)(hdrQ, {
                  title: `電力成本（overlay，預設關閉）`,
                  tip: `電力 ＝ 已連網 MW × 8760 × PUE × 電價 ÷ 1e9 × 期間係數（十億美元）；維護 ＝ 已連網 MW × 單價 ÷ 1000 × 期間係數。EBITDA 率已含電費，因此預設不另扣；若貢獻率改用未扣電費的毛數，才打開這個 overlay。`
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
                  tip: `信用損失 ＝ 收入 × 違約率 × (1−回收率)。${TXQ.concentration.title}；違約率是對此集中度的定價，不是預測。新債利率是瀑布「新增借款」的成本（${TXQ.newDebtRateNote}），存量債務的利息在「收支假設」頁單獨推算。`
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
                        children: [Y(d.leaseTail), `bn = 在帳 ${PERIODS[4]} `, LEASE_CASH_ON_BAL[4], ` × `, e.terminal.residualLeaseYears, ` 年（年報營業租賃加權剩餘租期 12 年）＋未起租租賃模型期後未付 `, Y(S.tail), `（總額 − 五期）。五期表外現金 `, Y(S.cashFive), ` + 尾 `, Y(S.tail), ` = `, Y(S.mapped), ` vs 承諾 `, Y(S.committed), `；在帳之後尚有 `, Y(LEASE_AFTER_FY30), `。`]
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
                  tip: `具名站點可編輯。公司不逐站揭露 MW；下列 MW 為公司公告的站點電力（口徑不明，÷1.2 換 MW-IT）。最後一列是殘差＝年度已連網 MW − 具名加總；殘差偏大是必然，代表容量預測多數不是由具名站點支撐。`
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
                  title: `站點租賃：公司揭露 vs 季報總額`,
                  tip: `季報不揭露逐站租金；自有站點占合約電力 >75%（公司說法），租賃以共置站點為主。下表列出具名站點的合約資訊（有揭露者），用來檢驗每 MW 年租金與模型租金路徑是否一致；未揭露者標「—」。`
                }), (0, $.jsx)(`div`, {
                  style: xstyQ.wrap,
                  children: (0, $.jsxs)(`table`, {
                    style: {
                      ...xstyQ.table,
                      minWidth: 860
                    },
                    children: [(0, $.jsx)(`thead`, {
                      children: (0, $.jsx)(`tr`, {
                        children: [`站點`, `契約 MW`, `合約總值 $bn`, `年期`, `年租金 $bn`, `每 MW 年租金 $m`, `季報對應`].map((e, t) => (0, $.jsx)(`th`, {
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
                          children: `—`
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
                    [`季報對帳`, null],
                    [`季報已揭露租賃承諾（在帳 ${Y(LEASE_FACTS.onBal, 2)}＋未起租 ${Y(LEASE_FACTS.notCommenced, 2)}）`, [LEASE_FACTS.onBal + LEASE_FACTS.notCommenced + LEASE_FACTS.singleCap, 0, 0, 0, 0], void 0, void 0, void 0, `未折現金額。`],
                    [`減：具名站點合約總值`, [-e.sites.filter(e => e.contract).reduce((e, t) => e + t.contract, 0), 0, 0, 0, 0]],
                    [`＝ 未具名站點（殘差）`, [LEASE_FACTS.onBal + LEASE_FACTS.notCommenced + LEASE_FACTS.singleCap - e.sites.filter(e => e.contract).reduce((e, t) => e + t.contract, 0), 0, 0, 0, 0], void 0, void 0, `tot`, `43 座以上資料中心中未具名者（託管商與其他房東）。殘差占比高是必然——公司不揭露逐站。`],
                    [`租金路徑檢驗（第三方租賃占 85%）`, null],
                    [`模型租金（在帳＋表外）`, d.years.map(e => e.lease)],
                    [`模型每 MW 年租金`, d.years.map(e => e.rentPerMW), void 0, void 0, `calc`, `＝模型租金 ÷ 期間長度 ÷ 平均 Accepted MW。FY26 欄只含下半年租金、MW 取全年平均，略低估。`, `US$m/MW`],
                    [`基準租金（MW × 市場基準 × 第三方占比 ${hA(LEASE_FACTS.share * 100, 0)}）`, d.years.map(e => e.rentBench), void 0, void 0, void 0, `市場基準＝具名站點加權每 MW 年租金；具名站點無租約金額時為 0（不適用）。`],
                    [`差額（基準 − 模型）`, d.years.map(e => e.rentBench - e.lease), void 0, void 0, `tot`, `正值＝模型租金可能低估的金額。本版只作檢驗，未改為 MW 驅動。`]
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
              }), n === 13 && (0, $.jsx)(QuarterTabQ, { d: d, p: f, st: e }), n === 12 && (0, $.jsxs)(`div`, {
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
                  tip: `把期前融資瀑布的結果按期整理成存量：期初現金＋營運收支淨額＋瀑布融資−新融資利息＝期末現金；既有債務依季報到期表遞減，加上期後新發可轉債、瀑布新債與高息債＝總債務。股數＝基礎股數（每年 SBC 稀釋 1%）＋累計新股。${PERIODS[0]} 欄的現金自評價日起算（模型期）。`,
                  w: 480
                }), (0, $.jsx)(BM, {
                  rows: [
                    [`現金`, null],
                    [`期初現金`, d.years.map((t, n) => n === 0 ? e.cash : d.years[n - 1].cum), void 0, void 0, void 0, `FY26 為 2026-06-30 現金（模型期起點）。`],
                    [`營運收支淨額（含期後股權／可轉債，不含瀑布）`, d.years.map(e => e.preFinGap), void 0, void 0, void 0, `＝營運來源＋期後股權／可轉債 − 所有支出（含還本、不含新融資利息）。負值即當期外部資金需求。`],
                    [`＋ 瀑布：新債（額度＋資產層）`, d.years.map(e => e.newDebt)],
                    [`＋ 瀑布：可轉債`, d.years.map(e => e.convNew)],
                    [`＋ 瀑布：高息債`, d.years.map(e => e.junk)],
                    [`＋ 瀑布：股權募資`, d.years.map(e => e.equity)],
                    [`− 新融資利息（新債＋可轉債＋高息債）`, d.years.map(e => -e.newDebtInt)],
                    [`期末現金`, d.years.map(e => e.cum), void 0, void 0, `tot`, `期前融資使期末現金不低於最低現金。`],
                    [`債務`, null],
                    [`既有債務期初（季報本金）`, d.years.map((t, n) => n === 0 ? (e.includeDebt ? DBT_P : DBT_P) : d.years[n - 1].existDebtEnd - CONV_P)],
                    [`− 排程還本`, d.years.map(e => -e.debtPay)],
                    [`既有債務期末`, d.years.map(e => e.existDebtEnd - CONV_P), void 0, void 0, `calc`],
                    [`＋ 期後新發可轉債`, d.years.map(() => CONV_P)],
                    [`＋ 瀑布新債餘額`, d.years.map(e => e.newDebtEnd)],
                    [`＋ 瀑布可轉債餘額`, d.years.map(e => e.convEnd)],
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
                    [`總債務 ÷ EBITDA（年化）`, d.years.map((e, t) => e.totalDebtEnd / Math.max(e.ebitdaPL / PERIOD_YEARS[t], .01)), void 0, void 0, void 0, `FY26 模型期 EBITDA 以半年 ×2 年化。`, `x`, 1],
                    [`租賃負債（期末）`, d.years.map(e => e.leaseLiab), void 0, void 0, void 0, `剩餘租金現值（折現率 ${hA(LLQ * 100, 1)}，10-K 加權平均）：在帳到期表＋未起租租約已起租部分（v0.2）。`],
                    [`調整後槓桿（(總債務＋租賃負債) ÷ (EBITDA＋租金)）`, d.years.map(e => e.adjLev), void 0, void 0, `tot`, `S&P 口徑近似；降評門檻 >${multTxt(e.debtCapBasis === `leaseAdj` ? e.debtEbitdaMax : 4.5)}×（二手轉述）。S&P 自身口徑另含全部未起租承諾與無條件採購義務，較本列高。`, `x`, 2],
                    [`距投資級上限的空間`, d.years.map(t => e.debtEbitdaMax - t.adjLev), void 0, void 0, void 0, `負值＝超過上限：需股權或失去投資級。`, `x`, 2]
                  ]
                })]
              }), n === 10 && (0, $.jsxs)(`div`, {
                className: `space-y-3`,
                children: [(0, $.jsx)(hdrQ, {
                  title: `既有債務（季報，${CALQ.valuationDate}）`,
                  tip: `有效利率為現金票息口徑（可轉債＝票息 ÷ 到期累積倍數，本金以到期累積本金計）。存量利息＝平均本金（依到期表遞減）× 加權有效利率 × 期間長度，另加期後新發可轉債利息與首期校準列。限制：假設利率組合不變；可轉債的折價攤銷（非現金）不計入。`
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
                        children: [`合計／加權`, ``, ``, hA(DBT_R * 100, 1), Y(DBT_P, 3), Y(DBT_I, 3), `＝季報本金 ${Y(LATEST_Q.debtPrincipal, 3)}`].map((e, t) => (0, $.jsx)(`td`, {
                          style: {
                            ...(t === 0 ? xstyQ.tdL : xstyQ.td),
                            background: `#eef2f7`,
                            fontWeight: 700
                          },
                          children: e
                        }, t))
                      }), (0, $.jsx)(`tr`, {
                        children: [`期後新發可轉債`, `追索`, `—`, hA(COMPANY_DATA.debt.convertible.coupon * 100, 3), Y(CONV_P, 3), Y(CONV_I, 3), `不在五期還本表（只計利息）`].map((e, t) => (0, $.jsx)(`td`, {
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
                    [`存量利息推導（本金依季報到期表遞減）`, null],
                    [`期初本金`, d.years.map(e => e.pBeg)],
                    [`排程還本`, [...DEBT_AMORT]],
                    [`期末本金`, d.years.map(e => e.pEnd)],
                    [`可轉債到期還本（債務處理）`, d.years.map(e => e.cvAmort)],
                    [`可轉債期末餘額（債務處理）`, d.years.map(e => e.cvEnd)],
                    [`可轉債票息（債務處理）`, d.years.map(e => e.cvInt)],
                    [`存量債務利息（含可轉債與首期校準）`, d.years.map(e => e.intStock), void 0, void 0, `tot`, `＝其他借款平均本金 × ${hA(DBT_R * 100, 1)} × 期間長度 ＋ 債務處理可轉債票息（原始本金 × 票息；到期當期計半期）＋ 首期校準 ${Y(DEFAULTS.intCal, 2)}。價內可轉債以若轉換法計，不計利息。`],
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
                          children: `季報（入帳）`
                        }), (0, $.jsx)(`th`, {
                          className: `py-2 font-medium`,
                          children: `本模型取捨`
                        })]
                      })
                    }), (0, $.jsx)(`tbody`, {
                      className: `font-mono tabular-nums`,
                      children: [
                        ...TXQ.callVsFiling,
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
                    children: `表外與契約性支出總表（季報，${CALQ.valuationDate}，US$bn）`
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
                          ...TXQ.commitments
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
                        children: `未起租 $${Y(LATEST_Q.offBalanceLease, 1)}bn 是契約，不是或有事項。`
                      }), `ASC 842 規定租賃在 commencement 才入表；已簽約、機房未交付 → 表外揭露。${TXQ.leaseNote}。`]
                    }), (0, $.jsx)(`p`, {
                      children: `客戶預付是合約負債，不是收入：預付現金在建置前先收，之後隨服務提供認列為營收（非現金）。${TXQ.prepayRiskNote}`
                    }), (0, $.jsxs)(`p`, {
                      children: [(0, $.jsx)(`span`, {
                        className: `font-medium text-fg`,
                        children: `可轉債與「營運缺口」的關係：`
                      }), `可轉債到期本金（含累積）列為契約性還本；是否轉換取決於到期時股價與轉換價。模型把還本列為契約性用途、把新融資列為「未籌」，避免「借新還舊自動成立」的隱含假設。資產擔保融資額度是第一順位填補工具，未預先扣減缺口。`]
                    })]
                  })]
                }), (0, $.jsxs)(Jj, {
                  className: `border-border bg-surface p-4`,
                  children: [(0, $.jsx)(`h3`, {
                    className: `font-display text-base font-semibold`,
                    children: `${TXQ.consistencyTitle}`
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
                            children: `新聞稿／法說`
                          }), (0, $.jsx)(`th`, {
                            className: `py-2 pr-3 font-medium`,
                            children: `季報／期後公告`
                          }), (0, $.jsx)(`th`, {
                            className: `py-2 font-medium`,
                            children: `判定`
                          })]
                        })
                      }), (0, $.jsx)(`tbody`, {
                        children: [
                          ...TXQ.consistency
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
                }), ...TXQ.sources.map(([tg, tx], k) => (0, $.jsx)(VM, { tag: tg, text: tx }, `s${k}`)), (0, $.jsx)(VM, {
                  tag: `Verified`,
                  text: consSourceTxtQ().peers
                }), (0, $.jsx)(VM, {
                  tag: CONSENSUS.priceTarget.tag,
                  text: consSourceTxtQ().targets
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
            children: `附錄｜資金橋（模型期：自 ${CALQ.valuationDate} 起；${PERIODS[0]} 全年口徑見「各期收支」頁底）`
          }), (0, $.jsxs)(`p`, {
            className: `mt-2 font-mono text-sm tabular-nums leading-relaxed`,
            children: [`期初 `, Y(y.open), ` + 營運缺口 `, Y(y.operatingGap), ` + 股權／可轉債 `, Y(y.atm), y.facility ? ` + 瀑布融資（新債＋股權）${Y(y.facility)}` : ``, y.debtPay ? ` − 排程還本 ${Y(y.debtPay)}` : ``, ` = 期末 `, Y(y.end), y.ok ? `` : `（對不上）`]
          }), (0, $.jsxs)(`p`, {
            className: `mt-2 text-xs leading-relaxed text-muted`,
            children: [`產能橋：排程 `, Y(b.scheduled), ` − 瓶頸 `, Y(b.bottleneck), ` − 信用損失 `, Y(b.credit), ` = 收現 `, Y(b.collected), `；新簽約現金 `, Y(d.totals.newCash), `（新簽約收入 `, Y(d.totals.newRev), `）。營收由 MW 驅動，RPO 只作對照。 RPO 滾動：評價日 `, Y(x.begin), ` + 期後新增 `, Y(x.q1Net), ` = `, Y(x.q1End), `；五期排程 `, Y(x.convert), `，2030 後仍未認列 `, Y(x.remaining), `。 支出五層：毛 CapEx `, Y(d.totals.gross), `（客戶預付後 `, Y(d.totals.cashCapex), `）＋ 租金 `, Y(d.totals.lease), `（在帳 `, Y(S.onFive), ` + 表外 `, Y(S.cashFive), `）＋ 利息 `, Y(d.totals.interest), `（其中瀑布新債 `, Y(d.totals.overdraft), `）＋ JV／投資 `, Y(d.totals.div), ` ＋ 排程還本 `, Y(d.totals.debtPay), `。表外現金租金五期 `, Y(S.cashFive), ` + 尾 `, Y(S.tail), ` = `, Y(S.mapped), `，對承諾 ${Y(S.committed, 1)} 差 `, Y(S.gap), `。`]
          })]
        }), (0, $.jsxs)(Jj, {
          className: `border-accent/20 bg-accent/5 p-4`,
          children: [(0, $.jsx)(`div`, {
            className: `text-xs font-medium text-muted`,
            children: SOURCE_ORDER_NOTE
          }), (0, $.jsx)(`div`, {
            className: `mt-3 grid grid-cols-2 gap-2 sm:grid-cols-3 lg:grid-cols-6`,
            children: [
              [`${CALQ.filedQLabel} 營收`, `${Y(LATEST_Q.revenue,3)}bn +${hA(LATEST_Q.yoy*100,0)} · 年初至今 ${Y(LATEST_Q.h1Revenue,3)}`],
              [`RPO 桶`, `${Y(LATEST_Q.rpo,1)}bn · ${COMPANY_DATA.rpo.split.map(x => hA(x*100,0)).join(`／`)}（${COMPANY_DATA.rpo.bucketLabels.join(`／`)}）· 只作對照`],
              [`CapEx`, `年初至今 ${Y(LATEST_Q.capexH1,2)}（現金）· 全年 ${CALL_FACTS.capexLo}–${CALL_FACTS.capexHi}（${TXQ.capexGuideSource}）`],
              [`債務本金`, `${Y(LATEST_Q.debtPrincipal,2)}bn · 五期攤還 ${Y(DEBT_AMORT.reduce((e,t)=>e+t,0),2)}`],
              [`客戶預付`, `合約負債 ${Y(LATEST_Q.deferredTotal,2)}bn · 年初至今淨增 ${Y(LATEST_Q.deferredIn,2)}`],
              [`租賃`, `在帳未折現 ${Y(LATEST_Q.onBalanceUndiscounted,2)} · 未起租 ${Y(LATEST_Q.offBalanceLease,1)}`]
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
              ...TXQ.headerCards,
              [`期後`, `${CALL_FACTS.postQShortDated}`],
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
