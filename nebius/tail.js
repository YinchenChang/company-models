
function tipQ({
  t: e,
  w: t = 360
}) {
  let [n, r] = (0, v.useState)(!1),
      [i, a] = (0, v.useState)({x:0,y:0}),
      o = (0, v.useRef)(null),
      s = () => {
        if(o.current){
          let b = o.current.getBoundingClientRect(),
              px = b.left,
              py = b.bottom + 6,
              vw = window.innerWidth,
              vh = window.innerHeight;
          if(px + t > vw) px = Math.max(8, vw - t - 8);
          if(py + 200 > vh) py = Math.max(8, b.top - 206);
          a({x:px,y:py})
        }
        r(!0)
      };
  return (0, $.jsxs)(`span`, {
    style: {
      display: `inline-block`
    },
    children: [(0, $.jsx)(`span`, {
      ref: o,
      onMouseEnter: s,
      onMouseLeave: () => r(!1),
      onClick: e => {
        e.preventDefault(), e.stopPropagation(), r(e => !e)
      },
      style: {
        cursor: `help`,
        marginLeft: 4,
        fontSize: 11,
        lineHeight: `11px`,
        color: `#0f5c61`,
        borderBottom: `1px dotted #0f5c61`,
        userSelect: `none`
      },
      children: `ⓘ`
    }), n ? (0, $.jsx)(`span`, {
      style: {
        position: `fixed`,
        zIndex: 9999,
        left: i.x,
        top: i.y,
        width: t,
        maxWidth: `78vw`,
        background: `#fff`,
        color: `#1e293b`,
        padding: `8px 10px`,
        borderRadius: 6,
        fontSize: 11,
        lineHeight: 1.7,
        boxShadow: `0 4px 16px rgba(0,0,0,.18)`,
        border: `1px solid #d1d5db`,
        whiteSpace: `normal`,
        textAlign: `left`,
        fontWeight: 400,
        pointerEvents: `none`
      },
      children: e
    }) : null]
  })
}

function hdrQ({
  title: e,
  tip: t,
  right: n
}) {
  return (0, $.jsxs)(`div`, {
    className: `flex flex-wrap items-baseline justify-between gap-2`,
    children: [(0, $.jsxs)(`h2`, {
      className: `font-display text-lg font-semibold`,
      children: [e, t ? (0, $.jsx)(tipQ, {
        t: t,
        w: 460
      }) : null]
    }), n || null]
  })
}
var xstyQ = {
  wrap: {
    width: `100%`,
    maxWidth: `100%`,
    overflowX: `auto`,
    border: `1px solid #c9c2b4`
  },
  table: {
    width: `100%`,
    minWidth: 760,
    borderCollapse: `collapse`,
    fontSize: 12,
    fontFamily: `ui-monospace, SFMono-Regular, Menlo, monospace`
  },
  th: {
    background: `#1f3864`,
    color: `#fff`,
    fontWeight: 600,
    padding: `6px 8px`,
    textAlign: `right`,
    border: `1px solid #16294a`,
    whiteSpace: `nowrap`,
    fontFamily: `inherit`
  },
  thL: {
    background: `#1f3864`,
    color: `#fff`,
    fontWeight: 600,
    padding: `6px 8px`,
    textAlign: `left`,
    border: `1px solid #16294a`,
    position: `sticky`,
    left: 0,
    zIndex: 3,
    minWidth: 220
  },
  sec: {
    background: `#d9e2f3`,
    fontWeight: 700,
    padding: `4px 8px`,
    border: `1px solid #c9c2b4`,
    textAlign: `left`,
    fontSize: 11.5
  },
  tdL: {
    padding: `3px 8px`,
    border: `1px solid #ddd6c8`,
    textAlign: `left`,
    position: `sticky`,
    left: 0,
    zIndex: 2,
    fontFamily: `ui-sans-serif, system-ui, sans-serif`,
    whiteSpace: `nowrap`
  },
  td: {
    padding: `3px 8px`,
    border: `1px solid #ddd6c8`,
    textAlign: `right`,
    whiteSpace: `nowrap`
  },
  unit: {
    padding: `3px 6px`,
    border: `1px solid #ddd6c8`,
    textAlign: `left`,
    color: `#6b7280`,
    fontSize: 10.5,
    whiteSpace: `nowrap`,
    fontFamily: `ui-sans-serif, system-ui, sans-serif`
  }
};

function nfmtQ(e, t = 1) {
  return Number.isFinite(e) ? e < 0 ? `(${Math.abs(e).toLocaleString(`zh-TW`,{minimumFractionDigits:t,maximumFractionDigits:t})})` : e === 0 ? `-` : e.toLocaleString(`zh-TW`, {
    minimumFractionDigits: t,
    maximumFractionDigits: t
  }) : `—`
}

function BM({
  rows: e,
  step: t = .1,
  cols: n = PERIODS,
  unit: r = `US$bn`
}) {
  let i = (e, t) => t ? {
    ...xstyQ.tdL,
    background: t === `sec` ? `#d9e2f3` : t === `tot` ? `#eef2f7` : e % 2 ? `#fbfaf7` : `#fff`,
    fontWeight: t === `tot` || t === `sec` ? 700 : 400
  } : {
    ...xstyQ.tdL,
    background: e % 2 ? `#fbfaf7` : `#fff`
  };
  return (0, $.jsx)(`div`, {
    style: xstyQ.wrap,
    children: (0, $.jsxs)(`table`, {
      style: xstyQ.table,
      children: [(0, $.jsx)(`thead`, {
        children: (0, $.jsxs)(`tr`, {
          children: [(0, $.jsx)(`th`, {
            style: xstyQ.thL,
            children: `項目`
          }), (0, $.jsx)(`th`, {
            style: {
              ...xstyQ.th,
              textAlign: `left`
            },
            children: `單位`
          }), n.map(e => (0, $.jsx)(`th`, {
            style: xstyQ.th,
            children: e
          }, e))]
        })
      }), (0, $.jsx)(`tbody`, {
        children: e.map(([e, a, o, s, c, l, d, dp], u) => a === null ? (0, $.jsx)(`tr`, {
          children: (0, $.jsx)(`td`, {
            colSpan: n.length + 2,
            style: xstyQ.sec,
            children: e
          })
        }, `s${u}`) : (0, $.jsxs)(`tr`, {
          children: [(0, $.jsxs)(`td`, {
            style: i(u, c),
            children: [e, l ? (0, $.jsx)(tipQ, {
              t: l
            }) : null]
          }), (0, $.jsx)(`td`, {
            style: {
              ...xstyQ.unit,
              background: u % 2 ? `#fbfaf7` : `#fff`
            },
            children: d || r
          }), a.map((e, n) => (0, $.jsx)(`td`, {
            style: {
              ...xstyQ.td,
              background: c === `tot` ? `#eef2f7` : u % 2 ? `#fbfaf7` : `#fff`,
              fontWeight: c === `tot` ? 700 : 400,
              color: !o && e < 0 ? `#9f1239` : `#1f2937`,
              padding: o ? `1px 3px` : `3px 8px`
            },
            children: o ? (0, $.jsx)(IM, {
              value: e,
              onChange: e => o(n, e),
              step: s ?? (/^MW/.test(d || r) ? 1 : t) // v0.2 第 2 輪：MW 類欄位以整數調整
            }) : nfmtQ(e, dp ?? (/^MW/.test(d || r) ? 0 : 1)) // v0.2 第 2 輪：MW 類欄位一律整數顯示
          }, n))]
        }, `${e}-${u}`))
      })]
    })
  })
}

function VM({
  tag: e,
  text: t
}) {
  return (0, $.jsxs)(`div`, {
    className: `rounded-md border-l-4 border-accent bg-surface px-3 py-2`,
    children: [(0, $.jsx)(`span`, {
      className: `font-mono text-xs text-accent`,
      children: e
    }), (0, $.jsx)(`p`, {
      className: `mt-1 text-xs leading-relaxed text-fg`,
      children: t
    })]
  })
}(0, y.createRoot)(document.getElementById(`root`)).render((0, $.jsx)(v.StrictMode, {
  children: (0, $.jsx)(zM, {})
}));
function accQ({
  title: e,
  sum: t,
  open: n = !1,
  children: r
}) {
  let [i, a] = (0, v.useState)(n);
  return (0, $.jsxs)(`div`, {
    style: {
      border: `1px solid #ddd6c8`,
      borderRadius: 8,
      background: `#fff`
    },
    children: [(0, $.jsxs)(`button`, {
      onClick: () => a(e => !e),
      style: {
        width: `100%`,
        display: `flex`,
        justifyContent: `space-between`,
        alignItems: `center`,
        gap: 8,
        padding: `8px 10px`,
        background: i ? `#1f3864` : `#eef2f7`,
        color: i ? `#fff` : `#1f2937`,
        border: 0,
        borderRadius: i ? `8px 8px 0 0` : 8,
        cursor: `pointer`,
        textAlign: `left`
      },
      children: [(0, $.jsxs)(`span`, {
        style: {
          fontWeight: 700,
          fontSize: 13
        },
        children: [i ? `▾ ` : `▸ `, e]
      }), (0, $.jsx)(`span`, {
        style: {
          fontSize: 11,
          opacity: .85,
          fontFamily: `ui-monospace, SFMono-Regular, Menlo, monospace`,
          whiteSpace: `nowrap`
        },
        children: t
      })]
    }), i ? (0, $.jsx)(`div`, {
      className: `space-y-4`,
      style: {
        padding: `10px 10px 12px`
      },
      children: r
    }) : null]
  })
}
var FMODS = [{
    name: `輸入與假設`,
    items: [[0, `收支假設`], [4, `信用／利率`], [5, `終值`]]
  }, {
    name: `各期收支`,
    items: [[1, `各期收支`], [13, `季度追蹤`]]
  }, {
    name: `運營活動`,
    items: [[2, `收入／產能`], [3, `電力成本`], [6, `站點`]]
  }, {
    name: `資產與負債`,
    items: [[10, `既有債務`], [11, `新債與新股`]]
  }, {
    name: `分析與檢查`,
    items: [[7, `敏感性`], [8, `連動檢查`], [9, `來源與承諾`], [12, `版本紀錄`]]
  }],
  VMODS = [{
    name: `損益`,
    items: [[0, `損益簡表`]]
  }, {
    name: `評價`,
    items: [[1, `Comps`], [2, `DCF`], [3, `目標價`], [4, `反向 DCF`]]
  }, {
    name: `市場共識`,
    items: [[5, `共識對照`]]
  }];

function navQ({
  groups: e,
  cur: t,
  set: n
}) {
  let r = e.find(e => e.items.some(e => e[0] === t)) || e[0];
  return (0, $.jsxs)(`div`, {
    className: `space-y-1`,
    children: [(0, $.jsx)(`div`, {
      className: `w-full max-w-full overflow-x-auto`,
      children: (0, $.jsx)(`div`, {
        className: `flex w-max gap-1 rounded-lg bg-surface p-1`,
        children: e.map(e => (0, $.jsx)(`button`, {
          onClick: () => n(e.items[0][0]),
          className: Vj(`h-10 shrink-0 rounded-md px-4 text-sm font-semibold`, e === r ? `bg-ink text-accent-fg` : `text-muted hover:text-fg`),
          children: e.name
        }, e.name))
      })
    }), r.items.length > 1 ? (0, $.jsx)(`div`, {
      className: `w-full max-w-full overflow-x-auto`,
      children: (0, $.jsx)(`div`, {
        className: `flex w-max gap-1 rounded-lg p-1`,
        style: {
          background: `#eef2f7`
        },
        children: r.items.map(([e, r]) => (0, $.jsx)(`button`, {
          onClick: () => n(e),
          className: Vj(`h-9 shrink-0 rounded-md px-3 text-sm font-medium`, t === e ? `bg-card text-fg shadow-card` : `text-muted hover:text-fg`),
          children: r
        }, r))
      })
    }) : null]
  })
}

var VLOG = [["v0.1", "10-07", "由 CRWV v4.5 模板建立；MW × 每 MW 收入主軸、預付款、融資瀑布、可轉債稀釋、其他事業與持股", "$61.7", "首版（基準）；保守 $63.8、積極 $154.6，三者皆賣出"], ["v0.2", "10-08", "工作單 v0.1d：標題改為 Nebius 專屬「預付款，能把 Backlog 變成現金嗎？」與動態副標；總結簡報改為 8 頁故事線＋附錄，一頁摘要移到最後；CRWV 論點句改寫；第 2 輪：MW 輸入改整數、DCF 常態化 FCF ≤ 0 不再判失效（終值以 0 計、保留權重）", "$61.7", "基準不變；保守 $63.8 → $35.2（(d) 方法：DCF 失效口徑）、積極 $154.6 → $154.5（(c) 假設：MW 整數）；時間推移、實際數皆 0"]];
