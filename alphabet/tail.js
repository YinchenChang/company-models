
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
              step: s ?? t
            }) : nfmtQ(e, dp ?? 1)
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

var VLOG = [["v0.1", "10-09", "由 Oracle v0.2 模板建立（CRWV v1.0–v4.5、Nebius v0.1–v0.2、Oracle v0.1–v0.2 的歷史見各自資料夾）；MAG 共用引擎：對外 AI 雲端 MW × Tokenomics 錨 × k、非 AI 雲端殘差、非雲端分部、回購彈性瀑布、持股、AI 增量報酬（含 C15 一致性檢查與讀法 2 並列）", "$150.5", "首版（基準）；保守 $174.8、積極 $130.3（三情境皆賣出）；讀法 2（自用 AI 價值中性）基準 $218.2；v0.1 交付前修訂（10-09，查核後）：一頁摘要加評價口徑敏感度（WACC 9／10／11%、ERP 4／6%、非 AI 自身倍數、AI 15×、組合；評等翻轉點）與終值占比規則句、自研晶片係數雙邊敏感度（TPU 收入與資本支出同乘 0.7）、Services 四線與 TPU 銷售起始 EBITDA 率標 [Assumed]、無形資產來源更正（數字不變）"], ["v0.2", "10-11", "接 Tokenomics v5.31（master f16f161）：快照改 v5.31（GB300 機架價格 −14%、IT 維護兩段機齡約 −50% → IF_HoldEcon GB300 −14%、其他世代 −3～−4%），名稱加四層瀑布 IFW_RevGWFleet_100／_Util／_Prod／_Life 與 IFC_Use；上限檢查改四層瀑布（利用率層示警 50%、生產折減層硬上限 100%）；C10 基準速度重解 1,069.4 → 1,246.1 MW-IT／年（積極＝1.5 × 基準）；引擎同步 amazon v0.2", "$145.9", "−$4.68（基準）＝(a) 收入錨 IF_HoldEcon +0.88＋(b) 每 MW 成本 +0.57＋(c) C10 速度重解 −5.81＋(d) 其他（Alphabet 層級費用隨營收重算）−0.31；保守 $174.8 → $175.8、積極 $130.3 → $122.0（皆賣出）；讀法 2 基準 $218.2 → $218.5"],
["v0.2.1", "10-22", "Tokenomics 同步 v19.0（master 1916a06）：快照換 v19.0（驅動名稱的值與 v5.31 全同，只有上限檢查層 IF_RevGWFleet／IFW_RevGWFleet_* 變動）；k 改規格（證據價格 ÷ 目前快照同世代成本）：k_長約 0.75 → 0.783、k_現貨 1.76 → 1.828、GCP 牌價 3.027 → 3.149（成本分母下降，價格不變）；上限檢查加第五列（每 MW 年收入 ÷ 模型商每 MW 實收 P×Q，無門檻）；C10 速度 1,246.1 → 1,247.9（積極 1,869.1 → 1,871.9）；Alphabet 層級費用線依新營收重算（占比 4.83% 不變）", "$146.1", "+$0.25（基準讀法 1）＝(1) 換快照 0.00＋(2) k 隨成本重算 +0.32＋(3) C10 重解 −0.06＋(4) 層級費用線 −0.01＋(5) 上限檢查 0.00；讀法 2 +0.33（$218.8）；保守 $175.1、積極 $122.9"],
["v0.2.2", "10-25", "Tokenomics 同步 v21.4（master 3ac70ac）：快照換 v21.4；v19.0 → v21.4 只有 VR200 低成本端（機架價低端 7.8 → 7.5）與上限檢查層（IF_RevGWFleet／IFW_RevGWFleet_* +0.4%～+0.6%、L1_RevGW_PQ_VR200 +1.49%）變動，基準與高成本欄全部不變；k 隨成本重算的結果與 v0.2.1 相同（k_長約 0.783／k_現貨 1.828）；自研晶片 NonNV 名稱登錄（IF_NonNV_Out_TPU／IF_NonNV_Hold_TPU，TPU 仍維持同期 NVIDIA 世代 × 1.0，不切換）；上限第五列說明改為 Q 與 GW 路線 token 帳同構；C10 速度不變（1,247.9；積極 1,871.9 MW-IT／年），層級費用線不變", "$146.1", "±$0.00（基準讀法 1）＝(1) 換快照 0.00＋(2) k 隨成本重算 0.00＋(3) C10 重解 0.00＋(4) 層級費用線 0.00＋(5) 上限檢查 0.00；讀法 2 $218.8；保守 $175.1、積極 $122.9（皆賣出）"]];
