
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
    items: [[2, `收入／產能`], [3, `電力成本`], [6, `站點`], ...(PMW_ONQ && FLEETQ ? [[14, `每 MW 經濟性`]] : [])] // W2：新方法時才顯示
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

var VLOG = [["v1.0", "09-18", "Oracle 模板改寫為 CoreWeave：2H26–FY30、10-Q／法說／8-K 數據", "$19", "基準當時為 8 GW"], ["v1.1", "09-18", "Comps（NBIS／IREN／APLD／CORZ）、CDS 720bps、未動用額度開關", "$19", "無變動"], ["v1.2", "09-18", "同業 EV 改由 SEC 申報自行計算；新增含租賃 EV", "$19", "無變動"], ["v1.3", "09-18", "回應 Grok 比較：Helios 526 MW、NOL 遞延稅、服務營收下修、2H26 利息 1.8、可轉債淨額 3.2；Excel 版", "$16", "NOL 與服務營收下修"], ["v1.4", "09-18", "FY26 改完整年度（1H 實際＋下半年模型）；Excel 式排版；浮動說明；扣缺口列以 0 為下限", "$16", "無變動"], ["v1.5", "09-18", "回應 Andy 反饋：站點租賃對帳、債務明細、JV 拆分、CapEx 與 MW 連動、刪除受產能約束列", "$7.7", "CapEx 連動使模型期 CapEx +15.6bn"], ["v1.6", "09-18", "HTML／Excel 一致性修正 7 處；Comps 改 EV/EBITDA 與 P/S", "$7.7", "無變動"], ["v1.7", "09-18", "權重移至 EV/EBITDA（45／55）；情境改為擴張力道 3／5／8 GW；車隊折舊；GPU 汰換", "$21.3", "基準改為 5 GW；情境定義改變，與前版不可直接比較"], ["v1.8", "09-18", "情境改為 4.2／5.6／8 GW", "$19.3", "基準改為 5.6 GW"], ["v1.9", "09-21", "移除 EV/Sales；現價更新至 9/21 收盤 $85.43；ATM 價值 6.2", "$19.5", "股價更新"], ["v2.0", "09-22", "期前融資瀑布＋稀釋取代透支利息＋扣缺口（債務上限 0.4× backlog）", "$32.4", "股權以高於模型內在價值的價格發行"], ["v2.1", "09-22", "更正先前「重複計費」的說法（僅文字）", "$32.4", "無變動"], ["v2.2", "09-22", "Comps 改以 EV/Sales 為主，EV/EBITDA 列為參考", "$32.4", "無變動"], ["v2.3", "09-22", "EBITDA 率單一來源（59%→65%）；資金端現金與損益連動；服務營收統一；債務上限 1.0x", "$21.3", "債務上限依 2026 年實際融資組合校準為主因"], ["v2.4", "09-22", "股權吸收上限 20%＋高息債 12%；融資前缺口圖；EBITDAR 更名；標題與摘要改寫", "$21.3", "無變動"], ["v2.5", "09-22", "敏感性改以目標價為指標；附錄移至結論下方", "$21.3", "無變動"], ["v2.6–2.8", "09-23", "外部（Copilot）版面調整與每 MW 年收入敏感性", "$21.3", "於 v2.9 併入"], ["v2.9", "09-23", "Excel 依類別排序並可收合；原生併入 Copilot 改動", "$21.3", "無變動"], ["v3.0", "09-23", "DCF 失效條件；DCF 下限方式切換（0／選擇權）；HTML 左欄分類選單；日期與版本同步", "$21.3", "選擇權模式為 $26.3"], ["v3.1", "09-23", "模組化（輸入與假設／各期收支／運營活動／資產與負債／損益／評價）；新債與新股；Excel 分頁改名與拆分", "$21.3", "無變動"], ["v3.2", "09-23", "切換情境時保留手動調整的數字", "$21.3", "無變動"], ["v3.3", "09-23", "反向 DCF（收入 × 成本組合）；版本變更紀錄；每 MW 年收入／建置成本倍數", "$21.3", "倍數預設 100%，無變動"], ["v3.4", "09-24", "新增「總結」簡報頁（8 頁、即時連動、簡報模式、列印 16:9）；股權敏感度表加註融資組合與自動說明；更正融資瀑布說明的股權年度；修正 Excel 無法開啟（sheetPr 子元素順序，v3.3 起即存在）", "$21.3", "無變動"], ["v3.5", "09-24", "EV/EBITDA 錨定年度可選（FY27–FY30，FY28 起以 WACC 折回 FY27 末）；評價頁與總結頁新增「錨定年度 × 倍數」矩陣；總結頁新增評價方法頁（9 頁）", "$21.3", "預設仍錨定 FY27，無變動；FY29 × 6x 為 $47.5"], ["v3.6", "09-24", "EV/EBITDA 錨定年度預設改為 FY29（接近穩態利潤率，與 6x 穩態倍數一致）", "$47.5", "錨定 FY27→FY29：EV/EBITDA 腿 $38.7→$86.4；保守情境 $71.7（−16%）接近賣出門檻"], ["v4.0", "09-24", "結構重整：公司資料集中到 company.json（HTML 與 Excel 共用，Excel 75 項輸入改讀此檔）；資料常數與引擎主函式改為可讀名稱；數字與畫面與 v3.6 完全相同", "$47.5", "無變動（純結構）"], ["v4.1", "09-25", "目標價改為區間呈現：情境區間（保守與積極情境，保留手動調整）與方法區間（目前情境，EV/EBITDA 5–6x）；判斷句揭露與賣出門檻的餘裕；DCF 權重說明；刪除評價頁舊的下檔／上檔與操作區間；Excel 情境區間以模擬運算表計算（活公式）", "$47.5", "點位無變動；情境區間 $32.1–$71.7、方法區間 $30.8–$47.5"], ["v4.2", "09-25", "設定集中：同業 Comps（數字、名稱、備註、說明文字）與評等門檻移入 company.json（peers、methodology.rating）；Excel「輸入與假設」F 區新增「評等門檻」7 格，評等代碼、賣出門檻價、判斷句與檢查頁改為引用；評等規則說明改為活公式；數字與畫面與 v4.1 相同", "$47.5", "無變動（純結構）"], ["v4.3", "09-25", "一頁摘要＋市場共識對照：總結頁新增一頁摘要（原 9 頁改為附錄，共 10 頁）；損益與評價新增「市場共識」分頁，Excel 新增「摘要」分頁與「輸入與假設」I 區（逐筆附來源、擷取日期、標記）；移除 Street 參考欄位；現價更新至 9/24 收盤 $90.13（股權發行價同步）", "$47.9", "現價 $85.43→$90.13：發行價提高、新股減少，基準 $47.5→$47.9；保守 −16%→−20%，距賣出門檻 $0.9→$4.9"], ["v4.4", "09-26", "季度層：資金模型新增「季度追蹤」分頁，Excel 新增「季度追蹤」分頁與「輸入與假設」J 區（2026 Q3–2027 Q4，由年度模型拆分、只用於追蹤；實際數輸入格預設空白）；差距超過門檻（營收等 5%、利潤類利潤率 2pt）或落在指引區間外的項目附差異原因（觀點／拆法／已知限制；拆法以水準／分配拆解判定）；差距寫法：營收、CapEx 為比例，利潤類為金額差＋利潤率百分點；一頁摘要驗證點改為焦點季 3 個指標，方法註腳與最新分析師動作移至市場共識分頁", "$47.9", "無變動（季度層不影響年度數字與目標價）"], ["v4.5", "09-26", "期間滾動＋反向 DCF：company.json 新增 calendar（財年結束月份、最新已申報／已公布季度），期間、評價日（最新已申報季末）與折現年數由日曆推算（首期 0.25／0.5／0.75／1 年、5 月財年、無實際數皆可）；actual1H 改為 ytdActual；目標價時點改為「評價日＋12 個月」（2027-06-30），EV/EBITDA 腿折回、DCF 腿（0 截斷與選擇權）以 WACC 推到同一時點；反向 DCF 改由 Excel 求解；一頁摘要與各分頁第一屏的日期與期間文字依日曆產生", "$45.5", "基準 $47.9→$45.5、保守 $71.7→$69.7、積極 $32.6→$31.0，三者皆賣出。拆解：(a) 時間推移 0（評價日仍為 6/30）、(b) 實際數更新 0、(c) 假設變更 0、(d) 方法變更（目標價時點改為評價日＋12 個月、DCF 腿推到同一時點）−$2.4（保守 −$1.9、積極 −$1.7）"], ["v4.6", "10-08", "每 MW 經濟性改接 Tokenomics v5.26（YinchenChang/Tokenomics，快照版本固定）：每 MW 資本支出＝新增世代占比 × IF_CapexIT（不含廠房，廠房在租金）、營運成本改由下而上（電費、IT 維護、人員軟體、稅險 × 平均在役 MW＋管銷率 5.5%）、折舊年限 IF_DeprLifeIT；收入採備案（每 MW 年收入仍為輸入，另列隱含 GPU 小時價格與 Tokenomics 持有成本對照）；世代組合（期初 1.5 GW：Hopper 27%／GB200 29%／GB300 43%；新增 GB300 → VR200）；方法開關 methodology.perMw（全設舊方法＝v4.5）；新增「每MW經濟性」頁、一頁摘要每 MW 一句", "$29.7", "基準 $45.5→$29.7、保守 $69.7→$44.0、積極 $31.0→$19.9，三者皆賣出。拆解：(a) 時間推移 0、(b) 實際數更新 0、(c) 假設變更 0、(d) 方法變更 −$15.8（① 每 MW 資本支出 −$7.6、② 折舊年限 0、③ 營運成本由下而上 −$8.2、④ 收入 0（備案）、⑤ MW 口徑 0）；保守 −$25.7（−$16.6／0／−$9.2）、積極 −$11.0（−$9.1／0／−$1.9）"]];
