# v4.2：產生 README「company.json 欄位說明」的表格。意義、單位、換公司時的處理寫在下方；「目前數值」與「程式使用」直接讀 company.json 與程式碼。
# company.json 有欄位沒寫說明、或說明了不存在的欄位，都直接報錯——新增欄位時先在這裡補說明，再重新產生。
# 用法：python3 scripts/fields_doc.py --write（直接更新 README「company.json 欄位說明」的「填表慣例」起至各區表格結束）；
#       --check 只比對、不一致即失敗（verify.sh 步驟 0b）；不帶參數則印出。
import json, sys, re, os

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CO = json.load(open(os.path.join(REPO, 'company.json'), encoding='utf-8'))
CODE = ''.join(open(f'{REPO}/{f}', encoding='utf-8').read() for f in
               ['segA.js', 'segB.js', 'segC.js', 'segD.js', 'segE.js', 'tail.js', 'build_xlsx.py', 'cmp31.js'])
sys.path.insert(0, REPO); import calendar_q, vlog  # 5a：期間字樣與版本號自動帶入（滾動或升版後不必手改）
CAL = calendar_q.derive(CO)
VER = vlog.VLOG[-1][0]
PL = [f"{CAL['periods'][0]} {CAL['stubWord']}" if CAL['stubMonths'] < 12 else CAL['periods'][0]] + CAL['periods'][1:]  # FY26 下半年、FY27…
P5 = f"五期（首期模型部分＋4 個完整財年；目前為 {'、'.join(PL)}）"
M, C, K = '必改', '檢查', '可沿用'   # 換公司：必改＝公司特有資料；檢查＝判斷值，依新公司重新評估；可沿用＝口徑或方法，通常不必改

# (路徑, 意義, 單位, 換公司)
F = [
 ('meta.company', '公司名稱', '文字', M), ('meta.ticker', '股票代號', '文字', M),
 ('meta.updateDate', '資料更新日', '日期', M), ('meta.priceDate', '股價日期（現價的收盤日；畫面與 Excel 的現價日期都讀這格）', '日期', M),
 ('meta.consensusFile', '市場共識資料檔路徑（v4.3；只讀，由使用者查證後提供；建置時併入 HTML、Excel 讀同一檔）', '路徑', M),
 ('texts.fy0EquityNote', '首期股權／可轉債的組成說明（含年初至今與首期模型的金額；v4.5；可用期間佔位符 «P0»、«YTD»、«STUB»）', '文字', M),
 ('texts.sourceLine', '頁首的資料來源一行（例如最新 10-Q、法說、期後 8-K；v4.5）', '文字', M),
 ('texts.cashTaxNote', '年初至今現金稅的說明（損益與評價頁；v4.5）', '文字', M),
 ('texts.thesis', '模型命題一句話（HTML 標題列與檔案說明；v0.1b）', '文字', M),
 ('texts.guideLine', '公司年度指引的一句摘要（損益頁、檢查與來源頁；v0.1b）', '文字', M),
 ('texts.taxNote', '稅率的來源說明（v0.1b）', '文字', M),
 ('texts.revMwCompare', '每 MW 年收入的公司對照值說明（只作對照；v0.1b）', '文字', M),
 ('texts.rpoNote', 'RPO 桶的說明（桶界、後段假設、口徑；v0.1b）', '文字', M),
 ('texts.leaseNote', '租賃結構的一句說明（自有或租賃為主；v0.1b）', '文字', M),
 ('texts.offBalanceLeaseTerm', '未起租租賃的起租時程與期限（季報揭露；v0.1b）', '文字', M),
 ('texts.leaseLiabNote', '租賃負債的折現率與期限說明（v0.1b）', '文字', M),
 ('texts.capexGuideSource', '資本支出指引的來源與可信度（v0.1b）', '文字', M),
 ('texts.prepayCheck', '客戶預付的公司說法與推得覆蓋比（檢查頁；v0.1b）', '文字', M),
 ('texts.mwFacts', '產能（MW）相關的公司揭露摘要（檢查頁；v0.1b）', '文字', M),
 ('texts.ebStartSource', '起始 EBITDA 率的來源或推導方式（v0.1b）', '文字', M),
 ('texts.costMwNote', '每 MW 建置成本的來源與口徑（v0.1b）', '文字', M),
 ('texts.ppeOpenNote', '期初 PP&E 基礎的來源或校準方式（v0.1b）', '文字', M),
 ('texts.facilityName', '未動用額度的名稱與條件（v0.1b）', '文字', M),
 ('texts.newDebtRateNote', '瀑布新債利率的來源（v0.1b）', '文字', M),
 ('texts.concentration', '客戶集中的標題與說明（title、detail；檢查頁與信用分頁；v0.1b）', '物件（文字）', M),
 ('texts.sharesNote', '評價股數在季末流通股之外的組成（v0.1b）', '文字', M),
 ('texts.netDebtNote', '淨負債輸入的口徑說明（v0.1b）', '文字', M),
 ('texts.otherRevNote', '非算力服務／其他事業營收的說明（v0.1b）', '文字', M),
 ('texts.rvRevNote', '反向 DCF 每 MW 年收入列的對照說明（v0.1b）', '文字', M),
 ('texts.rvCostNote', '反向 DCF 每 MW 建置成本列的對照說明（v0.1b）', '文字', M),
 ('texts.priceNote', '定價（續約價、新約價）的公司說法與讀法（v0.1b）', '文字', M),
 ('texts.prepayCoverNote', '客戶預付覆蓋比的來源（v0.1b）', '文字', M),
 ('texts.prepayOpenNote', '期初合約負債的口徑（v0.1b）', '文字', M),
 ('texts.ytdEquityNote', '年初至今股權募資的組成（v0.1b）', '文字', M),
 ('texts.prepayRiskNote', '客戶預付的風險與會計說明（v0.1b）', '文字', M),
 ('texts.consistencyTitle', '「公司說法 vs 季報」對照表的標題（v0.1b）', '文字', M),
 ('texts.callVsFiling', '來源頁「法說／季報／本模型取捨」對照表，一列一項：[項目, 前瞻說法, 季報, 本模型取捨]（v0.1b）', '清單', M),
 ('texts.commitments', '來源頁「表外與契約性支出總表」，一列一項：[項目, 金額, 時程, 入表？, 模型位置]（v0.1b）', '清單', M),
 ('texts.consistency', '來源頁「公司說法 vs 季報」一致性表，一列一項：[項目, 公司說法, 季報, 判定]（v0.1b）', '清單', M),
 ('texts.sources', '來源頁的來源清單，一列一筆：[標記, 內容]（HTML 與 Excel「來源」頁共用；v0.1b）', '清單', M),
 ('texts.headerCards', '資金模型頁首的公司揭露卡片，一列一張：[標題, 內容]（v0.1b）', '清單', M),
 ('texts.legacyMarginNote', '傳統事業 EBITDA 率的推導說明（v0.1b）', '文字', M),
 ('texts.mwYearEndNotes', '各年底主動電力的來源說明，以年份為鍵（Excel「輸入與假設」說明欄；5a）', '物件（文字）', M),
 ('meta.sourceOrderNote', '資料來源的先後與衝突時的取捨原則（畫面說明文字）', '文字', M),
 ('calendar.fiscalYearEndMonth', '財年結束月份（v4.5；CoreWeave 12、Oracle 5）', '月', M),
 ('calendar.latestQuarterFiled', '最新已申報（10-Q／10-K）的財季，格式 FYyyQn；驅動年度首期滾動與評價日（v4.5）', '文字', M),
 ('calendar.latestQuarterReported', '最新已公布（財報新聞稿）的財季；驅動季度層，年度首期不受影響（v4.5）', '文字', M),
 ('calendar.firstModelFY', '沒有已申報季度時（例如未上市公司）的首個模型財年，例如 FY26；有已申報季度時填 null（v4.5）', '文字', C),
 ('calendar.modelYears', '首期之後的完整財年數（固定期數滾動；Excel 欄位固定 5 期，所以為 4）（v4.5）', '年', K),
 ('calendar.targetHorizon', '目標價時點：firstFullYearEnd＝首期之後第一個完整財年末；valuationPlus12m＝評價日＋12 個月（v4.5）', '代碼', K),
 ('asOf', '滾動檢查：首期一次性金額與期初餘額所屬的已申報季度（鍵＝欄位路徑，清單定義在 calendar_q.py → ROLL_FIELDS）。每季 10-Q 後逐項更新數值，並把季度改為 calendar.latestQuarterFiled；缺漏或季度不符即建置失敗', '物件（季度）', M),
 ('ytdActual.throughQuarter', '年初至今實際數的截止財季，須等於 calendar.latestQuarterFiled（v4.5，原 actual1H）', '文字', M),
 ('ytdActual.months', '年初至今的月數，須等於日曆推算值（v4.5）', '月', M),
 ('ytdActual.label', '「年初至今實際數」的標題', '文字', M),
 ('ytdActual.jvSplit', 'JV 出資與策略投資的拆分（v4.5；說明文字與「JV 已付」讀此）', '物件（US$bn）', M),
 ('ytdActual.notes', '各欄位的逐列說明（來源、口徑、拆分；Excel「輸入與假設」G 區；v4.5 起隨資料一起更新）', '物件（文字）', M),
 ('ytdActual.cash1231', '上一年底現金', 'US$bn', M), ('ytdActual.revenue', '上半年營收', 'US$bn', M),
 ('ytdActual.capex', '上半年資本支出（認列口徑，含設備商融資）', 'US$bn', M),
 ('ytdActual.cashCapex', '上半年現金購置固定資產', 'US$bn', M), ('ytdActual.interest', '上半年利息費用', 'US$bn', M),
 ('ytdActual.leasePaid', '上半年租賃現金支付', 'US$bn', M), ('ytdActual.debtRepaid', '上半年還款', 'US$bn', M),
 ('ytdActual.borrow', '上半年借款', 'US$bn', M), ('ytdActual.equity', '上半年股權募資', 'US$bn', M),
 ('ytdActual.cappedCall', '上半年 capped call（可轉債配套避險）支出', 'US$bn', M),
 ('ytdActual.jv', '上半年合資（JV）與策略投資出資', 'US$bn', M), ('ytdActual.cfo', '上半年營運現金流', 'US$bn', M),
 ('ytdActual.prepay', '上半年客戶預付（遞延收入淨流入）', 'US$bn', M), ('ytdActual.da', '上半年折舊攤銷', 'US$bn', M),
 ('ytdActual.sbc', '上半年股份基礎薪酬', 'US$bn', M), ('ytdActual.opInc', '上半年 GAAP 營業損益', 'US$bn', M),
 ('ytdActual.ni', '上半年淨損益', 'US$bn', M), ('ytdActual.eps', '上半年 GAAP 每股盈餘', 'US$', M),
 ('ytdActual.ngEps', '上半年每股盈餘（加回股份基礎薪酬）', 'US$', M),
 ('ytdActual.dividends', '年初至今股利支付（普通股＋特別股；融資活動；v0.1b）', 'US$bn', M),
 ('ytdActual.adjEbitda', '上半年調整後 EBITDA（v4.3；只用於與市場共識比較 FY26，不進模型損益與評價）', 'US$bn', M),
 ('ytdActual.adjEbitdaMeta', '上述數字的 Q1／Q2 拆分、來源、標記與備註（建置時檢查 Q1＋Q2＝合計）', '物件', M),
 ('historicalPL', '歷年損益，一年一列（損益頁的歷史欄）。每列欄位：year 年度、revenue 營收、opInc GAAP 營業損益、ni 淨損益（以上 US$bn）、eps GAAP 每股盈餘、ngEps 加回股份基礎薪酬的每股盈餘（US$）、shares 加權流通股數（bn）、tag 來源標記', '清單', M),
 ('rpo.scheduledShare', '剩餘履約義務（RPO，已簽約未認列的營收）預計在模型期內認列的比例', '比例', M),
 ('rpo.bucketLabels', 'RPO 季報桶的名稱（與 rpo.split 對應；v0.1b）', '文字清單', M),
 ('rpo.split', 'RPO 季報桶的比例（季報揭露；v0.1b）', '比例清單', M),
 ('rpo.within36m', 'RPO 預計 36 個月內認列的比例（季報／法說；只作對照列；v0.1b）', '比例', M),
 ('rpo.bucketWeights', f'RPO 在{P5}各期認列的比例，合計等於上一欄', '比例清單', M),
 ('leases.onBalanceCash', f'已入帳租約在{P5}各期的現金租金', 'US$bn 清單', M),
 ('leases.afterFY30', '已入帳租約在模型期之後還要付的租金合計', 'US$bn', M),
 ('leases.facts.onBal', '已入帳租約未折現付款合計', 'US$bn', M),
 ('leases.facts.notCommenced', '已簽約但尚未起租的租約（表外）', 'US$bn', M),
 ('leases.liability.discRate', '租賃負債折現率（10-K 加權平均；v0.2）', '比例', M),
 ('leases.liability.tailYears', '在帳到期表模型期後尾端的平均分攤年數（租賃負債用；v0.2）', '年', C),
 ('leases.liability.note', '租賃負債口徑說明（v0.2）', '文字', M),
 ('leases.uncommenced.startQ', '未起租租賃自評價日後第幾季開始起租（0＝首期第一季）', '季', M),
 ('leases.uncommenced.quarters', '未起租租賃平均分攤起租的季數', '季', M),
 ('leases.uncommenced.termYears', '每筆未起租租賃的租期（直線付租）', '年', C),
 ('leases.uncommenced.termSens', '租期敏感度（報告用）', '年 清單', C),
 ('leases.uncommenced.note', '起租排程的來源與假設說明', '文字', M),
 ('leases.uncommenced.delayLink', '未起租租約起租隨建設延誤後移的比例（0–1；其餘照原時程；v0.2）', '比例', C),
 ('leases.uncommenced.delayLinkNote', 'delayLink 的依據說明（v0.2）', '文字', M),
 ('leases.facts.singleCap', '單一大型站點的租金上限（10-Q 揭露）', 'US$bn', M),
 ('leases.facts.share', '第三方租賃占機房取得的比例（用於租金基準檢驗）', '比例', C),
 ('leases.operatingPayments', '營業租賃到期表：五期各期，最後一格為之後合計（Excel 租賃頁）', 'US$bn 清單', M),
 ('leases.financePayments', '融資租賃到期表：同上格式', 'US$bn 清單', M),
 ('debt.amortization', f'既有債務在{P5}各期的排程還本', 'US$bn 清單', M),
 ('debt.amortAfterFY30', '模型期之後的還本合計', 'US$bn', M),
 ('debt.instruments', '既有債務逐筆明細，一筆一列：[名稱, 追索／非追索, 到期, 有效利率（比例）, 本金（US$bn）, 備註]。加總本金須等於 10-Q 本金合計', '清單', M),
 ('debt.convertible.principal', '期後新發行可轉債本金（不在五期還本表內，只計利息）', 'US$bn', M),
 ('debt.convertible.coupon', '該可轉債票面利率', '比例', M),
 ('debt.convertibles', '可轉債逐檔，一檔一列：[名稱, 原始本金（US$bn）, 票息（比例）, 到期 YYYY-MM, 到期累積倍數, 轉換價（US$）, 備註, 強制轉換（選填；true＝一律轉股、不計利息與還本、不列債務本金；Oracle v0.1b）]。有效轉換價＝轉換價 × 累積倍數；低於判斷價視為轉股（若轉換法），否則以到期累積本金計債務並付現金票息（v0.1b）', '清單', M),
 ('debt.convertibleBridge.exchangedAccreted', '評價日後以股換債註銷的舊債到期本金（季報本金與逐檔清單的調節項）', 'US$bn', M),
 ('debt.convertibleBridge.newIssuesAccreted', '評價日後新發可轉債的到期本金（季報本金與逐檔清單的調節項）', 'US$bn', M),
 ('debt.convertibleBridge.note', '調節說明與來源', '文字', M),
 ('scenarios.labels.low', '保守情境名稱（空格前的文字會當作情境簡稱）', '文字', M),
 ('scenarios.labels.base', '基準情境名稱', '文字', M), ('scenarios.labels.high', '積極情境名稱', '文字', M),
 ('scenarios.mwPath.connectedStart', '首期期末已連網 MW（三情境共用；v0.1b）', 'MW', M),
 ('scenarios.mwPath.contracted.low', f'保守情境：{P5}各期末的合約 MW 上限（已連網不得超過；v0.1b）', 'MW 清單', M),
 ('scenarios.mwPath.contracted.base', '基準情境：同上', 'MW 清單', M), ('scenarios.mwPath.contracted.high', '積極情境：同上', 'MW 清單', M),
 ('scenarios.mwPath.pace.low', '保守情境：併網速度（每年新增已連網 MW；已連網＝MIN(合約上限, 前期＋速度×期間長度)；v0.1b）', 'MW／年', C),
 ('scenarios.mwPath.pace.base', '基準情境：同上', 'MW／年', C), ('scenarios.mwPath.pace.high', '積極情境：同上', 'MW／年', C),
 ('scenarios.mwPath.note', '已連網 MW 路徑的說明（來源與口徑）', '文字', M),
 ('scenarios.revMW.low', '保守情境：每 MW 年收入，各期（Tokenomics 正向推導；不得用公司 ACV；v0.1b）', 'US$bn/MW 清單', M),
 ('scenarios.revMW.base', '基準情境：同上', 'US$bn/MW 清單', M), ('scenarios.revMW.high', '積極情境：同上', 'US$bn/MW 清單', M),
 ('scenarios.descriptions', '三情境的一句說明（Excel A 區；v0.1b）', '物件（文字）', M),
 ('scenarios.billableRatio.mode', '可計費 MW 的算法：ratio＝已連網 × 比例；converge＝期初以最新季實際營收年化 ÷ 每 MW 年收入校準，之後向已連網收斂（v0.1b）', '代碼', C),
 ('scenarios.billableRatio.openAnnualRevenue', 'converge 模式的校準營收：最新季實際營收（OCI 等 MW 驅動部分）× 4（v0.1b）', 'US$bn', M),
 ('scenarios.billableRatio.note', '可計費 MW 校準與收斂比例的說明（v0.1b）', '文字', M),
 ('scenarios.billableRatio.ratio', '各期比例：ratio 模式＝在役 ÷ 已連網；converge 模式＝期初校準值向已連網收斂的比例（三情境共用；v0.1b）', '比例清單', C),
 ('scenarios.mw31.low', '保守情境：模型期後一年（FY31）新增的 MW，用於 FY30 的預建支出', 'MW', C),
 ('scenarios.mw31.base', '基準情境：同上', 'MW', C), ('scenarios.mw31.high', '積極情境：同上', 'MW', C),
 ('scenarios.convCap.low', '保守情境：融資瀑布可轉債步驟每年新發行上限（0＝不新發；v0.1b）', 'US$bn／年', C),
 ('scenarios.convCap.base', '基準情境：同上', 'US$bn／年', C), ('scenarios.convCap.high', '積極情境：同上', 'US$bn／年', C),
 ('scenarios.delayMonths.low', '保守情境：建設延誤月數——計費 MW＝原可計費路徑往後平移此月數（期間長度線性內插）；GPU 資本支出與客戶出資照原時程，折舊自投入使用起算（v0.2）', '月', C),
 ('scenarios.delayMonths.base', '基準情境：同上', '月', C), ('scenarios.delayMonths.high', '積極情境：同上', '月', C),
 ('scenarios.delayMonths.note', '建設延誤月數的依據與說明（v0.2）', '文字', M),
 ('scenarios.capexTemplate.costMW', '每 MW 建置成本（GPU＋網路＋機房內裝），各期', 'US$m/MW 清單', C),
 ('scenarios.capexTemplate.div', 'JV 後續增資與策略投資，各期', 'US$bn 清單', C),
 ('legacy.capexV14', '舊版（v1.4）手動資本支出，只供畫面對照', 'US$bn 清單', K),
 ('legacy.interestV14', '舊版（v1.4）手動利息，只供畫面對照', 'US$bn 清單', K),
 ('defaults.scenario', '開啟時的預設情境（low／base／high）', '文字', C),
 ('defaults.revenueDriver', '營收驅動：mw＝平均在役 MW × 每 MW 年收入 × 利用率（新產能簽約率固定 100%，RPO 只作對照）；rpo＝CRWV 模板的 RPO 排程＋新簽約（v0.1b）', '代碼', C),
 ('defaults.lambda', '提前支出比例：次年才上線的 MW，其建置支出落在前一年的比例', '比例', C),
 ('defaults.mwYearEnd', '各年底主動電力，以年份為鍵（例如 "2025": 850）：首期期初取首期前一財年末；GPU 汰換批次＝各年新增 MW（5a；列入滾動檢查）', 'MW 物件', M),
 ('defaults.mw31', '模型期後一年新增 MW 的預設值（情境切換時改用 scenarios.mw31）', 'MW', C),
 ('defaults.capexFloorFY0', '首期所屬財年的全年資本支出下限（已下單的承諾，取公司指引下緣；首期＝下限 − 年初至今實際認列）', 'US$bn', M),
 ('defaults.gpuLife', 'GPU 經濟壽命（決定汰換時點與折舊）', '年', C),
 ('defaults.refreshSteady', '穩態汰換：true＝已連網 MW 不再增加的期間（觸頂後）及終值年，汰換 CapEx＝平均已連網 MW × 每 MW 建置成本 ÷ GPU 壽命 × 期間長度；false＝只有批次汰換（Oracle v0.1c）', '是／否', C),
 ('defaults.refreshNote', '穩態汰換的說明文字', '文字', K),
 ('defaults.revScale', '每 MW 年收入整體倍數（反向 DCF 與壓力測試用，1＝不調整）', '倍', K),
 ('defaults.capexScale', '每 MW 建置成本整體倍數（1＝不調整）', '倍', K),
 ('defaults.ebStart', '第一期 EBITDA 率', '比例', M), ('defaults.ebSteady', '最後一期（穩態）EBITDA 率；中間各期線性內插', '比例', C),
 ('defaults.ebitdaBasis', 'EBITDA 口徑：ebitdar＝EBITDA＝EBITDAR 率 × 營收 − 租金（租金為固定成本，EBITDAR 率三情境共用）；其他值＝三情境共用 EBITDA 率（模板）（Oracle v0.1c）', '代碼', C),
 ('defaults.ebitdarAdj', 'EBITDAR 率校準：基準情境 [首期, 末期] 租金 ÷ OCI 營收；EBITDAR 率＝ebStart／ebSteady＋此值（scripts/calib_ebitdar.js --write 產生，verify.sh 檢查）', '比例清單', M),
 ('defaults.ebitdarNote', 'EBITDAR 口徑的說明文字', '文字', K),
 ('defaults.services', '非算力服務營收（軟體、儲存等），各期', 'US$bn 清單', C),
 ('defaults.otherEbitda', '其他事業 EBITDA（負值＝燒錢），各期；同時進入損益 EBITDA 與營運來源（v0.1b）', 'US$bn 清單', M),
 ('defaults.legacyBiz.lines', '傳統事業各線，一線一列：key、label、fyBase 上一財年實際營收（US$bn）、ytd 年初至今實際營收、g0 起始年增率、gLT 長期年增率（自首期線性收斂到末期）（Oracle v0.1b）', '清單', M),
 ('defaults.legacyBiz.ebitdaMargin', '傳統事業 EBITDA 率，各期（Oracle v0.1b）', '比例清單', C),
 ('defaults.legacyBiz.note', '傳統事業輸入的來源與推導說明', '文字', M),
 ('defaults.cashTaxRate', '類現金流量的現金稅率：稅 ＝ 稅率 × MAX(0, 損益 EBITDA − 車隊 D&A − 存量利息)（v0.1b；虧損或 NOL 公司填 0）', '比例', C),
 ('defaults.delayPenalty', '延誤罰則／服務抵減：延誤期間應計費而未計費營收的比例，列為營業費用（預設 0＝未揭露；v0.2）', '比例', C),
 ('defaults.delayPenaltyNote', 'delayPenalty 的依據說明（v0.2）', '文字', M),
 ('defaults.debtCapBasis', '瀑布新債的上限基準：leaseAdj＝(總債務＋租賃負債) ≤ 倍數 ×(EBITDA＋租金)（年化；S&P 口徑近似的投資級上限，v0.2）；ebitda＝總債務 ≤ 倍數 × 當期 EBITDA（v0.1b）；backlog＝模板的債務／backlog', '代碼', C),
 ('defaults.debtEbitdaMax', '投資級上限倍數：leaseAdj＝調整後槓桿上限（S&P BBB- 降評門檻 4.5×，v0.2）；ebitda＝總債務 ÷ 當期 EBITDA 上限', '倍', C),
 ('defaults.dividend.perShareQ', '普通股每股每季股利（Oracle v0.1b；不發股利的公司刪除 dividend 區段）', 'US$', M),
 ('defaults.dividend.sharesBase', '股利的基礎股數（最新流通股；另加前期累計瀑布新股與已強制轉換特別股）', 'bn 股', M),
 ('defaults.dividend.preferred', '特別股股利，各期', 'US$bn 清單', M),
 ('defaults.dividend.note', '股利的來源與推導說明', '文字', M),
 ('defaults.otherEbitdaNote', '其他事業 EBITDA 的推導與來源說明', '文字', M),
 ('defaults.debtBacklog', '新債上限：總債務不超過 backlog 的倍數', '倍', C),
 ('defaults.ctrTerm', '新簽合約的平均年期（決定 backlog 補入量）', '年', C),
 ('defaults.minCash', '最低現金：每期融資後期末現金不低於此值', 'US$bn', C),
 ('defaults.eqPx', '新股發行參考價（預設＝現價）', 'US$', M),
 ('defaults.eqDisc', '新股發行折價', '比例', C),
 ('defaults.eqCapPct', '每年股權募資上限（占現市值）；輸入 9 以上視為無上限', '比例', C),
 ('defaults.eqCapShares', '股權年上限的股數基礎：現市值＝發行參考價 × 此股數（5a；評價日時點，列入滾動檢查）', 'bn', M),
 ('defaults.junkRate', '股權上限用完後的高息債利率', '比例', C),
 ('defaults.ppeOpen', '最新季末固定資產毛額', 'US$bn', M),
 ('defaults.jvCommit', '已承諾的 JV 出資餘額，各期', 'US$bn 清單', M),
 ('defaults.intCal', '第一期利息校準值（讓模型利息對上公司季度指引）', 'US$bn', M),
 ('defaults.rpoOpen', '最新季末 RPO', 'US$bn', M),
 ('defaults.rpoPendingAdd', '季末後新簽、尚未進 RPO 的承諾', 'US$bn', M),
 ('defaults.rp', 'RPO 在模型期內認列的比例（百分點；應與 rpo.scheduledShare 一致）', '%', M),
 ('defaults.cash', '最新季末現金（第一期期初現金）', 'US$bn', M),
 ('defaults.includeDebt', '是否依到期表攤還既有債務（true＝是；false＝假設全數再融資）', '是／否', K),
 ('defaults.includeAtm', '是否計入期後股權／可轉債募資', '是／否', K),
 ('defaults.atm', '期後股權／可轉債募資淨額（記在第一期）', 'US$bn', M),
 ('defaults.prepay.shareOfDeals', '有預付的合約比例（預付流入＝成長型 CapEx × 此比例 × 下一欄；覆蓋比已是整體口徑時填 1；v0.1b）', '比例', C),
 ('defaults.prepay.capexCover', '預付（客戶出資）占相關資本支出的比例', '比例', C),
 ('defaults.prepay.recogYears', '預付在合約期內的認列年數：依(期初合約負債＋本期累積利息)直線認列為營收（非現金）', '年', C),
 ('defaults.prepay.coverRefresh', '客戶出資覆蓋比是否也適用 GPU 汰換 CapEx（true＝預付流入＝(成長型＋汰換)× 覆蓋比；Oracle v0.1c）', '是／否', C),
 ('defaults.prepay.financingRate', '預付重大財務組成的隱含利率：合約負債以此利率累積非現金利息（期初餘額＋本期流入一半），認列時轉營收；0＝不計財務組成（Oracle v0.1b）', '比例', C),
 ('defaults.prepay.openBalance', '期初合約負債（客戶預付餘額；列入滾動檢查）', 'US$bn', M),
 ('defaults.prepay.note', '預付款區塊的說明（來源與口徑）', '文字', M),
 ('defaults.convIssue.coupon', '瀑布新發可轉債的票息（v0.1b）', '比例', C),
 ('defaults.convIssue.premium', '瀑布新發可轉債的轉換溢價（轉換價＝發行參考價 ×（1＋此值）；只用於潛在股數揭露）', '比例', C),
 ('defaults.convIssue.note', '可轉債步驟的說明（來源與口徑）', '文字', M),
 ('defaults.overlay', '電力／維護成本另計（預設關；EBITDA 率已含電費，開啟會重複扣除）', '是／否', K),
 ('defaults.cdsLink', 'CDS 利差是否傳入新債利率（預設關）', '是／否', K),
 ('defaults.cdsBaseBp', 'CDS 傳入新債利率的門檻：超過此值的部分才傳入（5a）', 'bps', C),
 ('defaults.cdsPassThrough', 'CDS 超額傳入新債利率的比例（每 1bp 傳入的 bp；5a）', '比例', C),
 ('defaults.linkSites', '具名站點的 MW 是否連動第一期產能下限', '是／否', K),
 ('defaults.linkLeaseTail', '舊模板遺留開關，目前程式未使用', '是／否', K),
 ('defaults.useAvgMw', '收入以平均在役 MW 計（true）或期末存量計（false）', '是／否', K),
 ('defaults.billableOpen', '最新季末可計費 MW（第一期期初）', 'MW', M),
 ('defaults.cds', '信用違約交換（CDS）中價', 'bps', M), ('defaults.cdsBid', 'CDS 買價', 'bps', M),
 ('defaults.cdsAsk', 'CDS 賣價', 'bps', M), ('defaults.cdsLo', 'CDS 近期區間下緣', 'bps', M),
 ('defaults.cdsHi', 'CDS 近期區間上緣', 'bps', M), ('defaults.cdsDate', 'CDS 報價日期與來源標記', '文字', M),
 ('defaults.useFacility', '融資時是否先動用未動用信用額度', '是／否', K),
 ('defaults.facility', '未動用信用額度', 'US$bn', M),
 ('defaults.m.accepted', '已驗收 MW 的預設路徑（實際依所選情境覆寫）', 'MW 清單', C),
 ('defaults.m.billable', '可計費 MW 的預設路徑（實際依情境與爬坡比例覆寫）', 'MW 清單', C),
 ('defaults.m.util', '利用率，各期', '% 清單', C), ('defaults.m.revMW', '每 MW 年收入，各期', 'US$bn/MW 清單', M),
 ('defaults.m.aiShare', 'AI 占比，各期（目前只做範圍檢查，未參與計算）', '% 清單', K),
 ('defaults.m.fill', '新產能簽約率：未被既有 RPO 占用的產能能賣出的比例', '% 清單', C),
 ('defaults.m.power', '電價（overlay 開啟時才用）', '$/MWh 清單', C), ('defaults.m.pue', '電力使用效率 PUE（overlay 用）', '倍 清單', C),
 ('defaults.m.maint', '維護成本（overlay 用）', 'US$m/MW 清單', C),
 ('defaults.m.defaultP', '客戶違約率', '% 清單', C), ('defaults.m.recovery', '違約回收率', '% 清單', C),
 ('defaults.m.rate', '新債利率', '% 清單', C),
 ('defaults.terminal.residual', '模型期末 GPU 殘值率', '%', C), ('defaults.terminal.rerent', '期末設備再出租率', '%', C),
 ('defaults.terminal.margin', '模型期後剩餘 RPO 的利潤率', '%', C),
 ('defaults.terminal.residualLeaseYears', '模型期後租約剩餘年數', '年', C),
 ('defaults.sites', '具名資料中心站點，一站一列。欄位：id 代碼、name 名稱、operator 房東／合作方、planned 契約 MW、energized 已通電 MW、accepted 已驗收 MW、billable 可計費 MW、contract 合約總值（US$bn，可無）、years 合約年期（可無）、status 狀態說明、next 下一里程碑、date 預計時間、confidence 信心（高／中／低）', '清單', M),
 ('valuation.price', '現價', 'US$', M), ('valuation.shares', '評價股數（含期後股權發行上限）', 'bn 股', M),
 ('valuation.atmSharesInValuation', '評價股數中「期後股權發行上限」的股數：期後股權／可轉債開關關閉時由評價股數扣回（v0.1b；CRWV 0.035、無此項的公司填 0）', 'bn 股', M),
 ('valuation.netDebt', '淨負債（不含可轉債：其他借款 − 現金，含期後已入帳的股權／可轉債募得淨額；可轉債依 debt.convertibles 另計）', 'US$bn', M), ('valuation.holdings', '持股清單，一筆一列：[名稱, 估值（100%，US$bn）, 持股比例, 備註]；價值＝估值 × 持股比例 ×（1 − holdingsDiscount），自淨負債扣除', '清單', M), ('valuation.holdingsDiscount', '持股折價（流動性、少數股權）', '比例', M), ('valuation.debtLike', '類債項目，一筆一列：[名稱, 金額（US$bn）, 備註]；加入淨負債', '清單', M), ('valuation.tax', '稅率', '比例', C),
 ('valuation.nol', '期初可扣抵虧損（NOL）', 'US$bn', M), ('valuation.wacc', '加權平均資金成本 WACC 手動覆蓋（null＝採 CAPM）', '比例或 null', C),
 ('valuation.capm.beta', 'CAPM β', '倍', C), ('valuation.capm.erp', 'CAPM 股權風險溢酬', '比例', C), ('valuation.capm.kdPretax', '稅前債務成本（市場邊際）', '比例', C), ('valuation.capm.betaSens', 'β 敏感度（報告用）', '倍 清單', C), ('valuation.capm.note', 'WACC 公式與來源說明', '文字', M),
 ('valuation.nolUsePct', 'NOL 每年可抵用上限占應稅所得的比例（美國 80%；依公司稅籍調整；5a）', '比例', C),
 ('valuation.wcPctOfRevGrowth', '營運資金變動占營收增量的比例（DCF 自由現金流；5a）', '比例', C),
 ('valuation.g', '永續成長率', '比例', C), ('valuation.sbc', '年度股份基礎薪酬', 'US$bn', M),
 ('valuation.maintRatio', '終值的維持性資本支出占折舊比例', '比例', C),
 ('valuation.tvBasis', '終值基準：ufcf＝末期 UFCF（含穩態汰換 CapEx）；其他值＝常態化 FCF（EBIT ×(1−稅)＋D&A × (1 − 維持比率)）（Oracle v0.1c）', '代碼', C),
 ('valuation.evEbitda', 'EV/EBITDA 倍數（分部加總的 OCI／算力部分）', '倍', C), ('valuation.evYear', 'EV/EBITDA 錨定年度（1＝模型第 2 期 … 4＝第 5 期）', '年度代碼', C),
 ('valuation.legacyEvEbitda', '傳統事業 EV/EBITDA 手動覆蓋（null＝軟體同業中位數）', '倍或 null', C), ('valuation.legacyEvEbitdaNote', '分部加總說明', '文字', M),
 ('valuation.dcfMode', 'DCF 股權為負時的處理：zero＝0 截斷、option＝選擇權法', '文字', K),
 ('valuation.sigma', '企業價值波動率（選擇權法用）', '比例', C), ('valuation.rf', '無風險利率（CAPM 與選擇權法用）', '比例', C),
 ('methodology.blendWeights.dcf', '加權目標價中 DCF 的權重', '比例', K),
 ('methodology.blendWeights.pe', '加權目標價中 EV/EBITDA 的權重（兩者合計 1）', '比例', K),
 ('methodology.rangeMultiples', '方法區間的 EV/EBITDA 倍數下端與上端', '倍 清單', C),
 ('methodology.consensusGapTol', '與市場共識比較的判斷句門檻：營收、EBITDA、CapEx 差距絕對值超過此比例即視為分歧（v4.3）', '比例', K),
 ('methodology.profitMarginTolPt', '利潤類（調整後 EBITDA、調整後營業利益）須附差異原因的門檻：利潤率差超過此百分點（寫成比例，0.02＝2pt）；營收、CapEx、淨負債、MW 仍用 consensusGapTol 的比例門檻；不影響分歧判斷句（v4.4）', '比例', K),
 ('methodology.rating.buyUpsideMin', '買進：空間（加權目標價 ÷ 現價 − 1）至少要達到的值', '比例', K),
 ('methodology.rating.buyTvShareMax', '買進：終值占企業價值須低於此值', '比例', K),
 ('methodology.rating.sellUpsideMax', '賣出：空間小於或等於此值即賣出；賣出門檻價＝現價 ×（1＋此值）', '比例（負數）', K),
 ('methodology.rating.equityRaiseMaxMult', '五期股權募資超過「現市值 × 此倍數」即禁止買進', '倍', K),
 ('methodology.rating.sellUpsideMaxIfEquityOver', '股權募資超標時，空間小於或等於此值即賣出', '比例（負數）', K),
 ('methodology.rating.sellMarginAlert', '任一情境與賣出門檻的距離小於「現價 × 此值」時，判斷句另外揭露', '比例', K),
 ('methodology.rating.tvShareWarn', '終值占企業價值超過此值時提出警示（融資說明、檢查頁）', '比例', K),
 ('methodology.checks.capexPerMwBand', '檢查頁：模型期 CapEx 強度（每 MW 百萬美元）的合理區間下端與上端（5a）', 'US$m/MW 清單', C),
 ('methodology.checks.leaseVsCommitMin', '檢查頁：表外租金路徑 ÷ 已承諾租約至少要達到的倍數（5a）', '倍', C),
 ('methodology.checks.unsignedRevShareMax', '檢查頁：後段年度依賴未簽約收入的比例上限（5a）', '比例', C),
 ('methodology.checks.siteRentGapMax', '檢查頁：站點租賃五期租金可能低估的金額上限（5a）', 'US$bn', C),
 ('methodology.checks.rentVsBenchMin', '檢查頁：模型每 MW 年租金至少要達到「市場基準 × 第三方占比」的比例（5a）', '比例', C),
 ('peers.priceDate', '同業市值的收盤日', '日期', M), ('peers.priceSource', '同業市值來源', '文字', M),
 ('peers.list', '同業，一家一列（HTML Comps 分頁與 Excel「可比公司」頁共用）。欄位：ticker 代號、name 公司名、labelHtml／labelXlsx 兩邊顯示的名稱、roleHtml 定位說明（HTML）、mkt 市值、netDebt 淨負債、rev 近十二個月營收、opl 營業租賃負債、opInc GAAP 營業損益、da 折舊攤銷（以上 US$bn）、asOf 資料期、noteHtml／noteXlsx 兩邊的備註。EV＝市值＋淨負債、EBITDA＝營業損益＋折舊攤銷，由程式計算', '清單', M),
 ('peers.software', '軟體同業 NTM EV/EBITDA，一家一列（ticker、name、ntmEvEbitda、ref＝事實總帳 id）；中位數為傳統事業倍數', '清單', M),
 ('peers.softwareNote', '軟體同業倍數的來源說明', '文字', M),
 ('peers.textHtml.headerTip', 'HTML Comps 表標題的浮動說明', '文字', M),
 ('peers.textHtml.readingTip', 'HTML「讀法」段落的浮動說明（折價來源）', '文字', M),
 ('peers.textHtml.caveat', 'HTML Comps 表下方的「口徑與限制」', '文字', M),
 ('peers.textXlsx.subtitle', 'Excel「可比公司」頁第 2 列說明', '文字', M),
 ('peers.textXlsx.notes', 'Excel「可比公司」頁表下方的讀法說明（每句一列）', '文字清單', M),
 # v4.4：季度層與差異原因
 ('quarterly._note', '季度層的說明文字（不進程式）', '文字', K),
 ('quarterly.quarters', '追蹤的季度，一季一列：key 季別代碼（與共識檔 quarterlyEstimates 的鍵相同，例如 2026Q3）、label 顯示名稱、period 所屬模型期（0＝第 1 期…）、reportNote 財報日說明（選填）。季度加總必須等於所屬模型期的年度數字', '清單', M),
 ('quarterly.periodNames', '各所屬模型期的顯示名稱（依期別順序，例如 2H26、FY27）', '文字清單', M),
 ('quarterly.focus', '焦點季（「季度追蹤」第一屏與一頁摘要驗證點顯示的季度；財報後改為下一季）', '季別代碼', M),
 ('quarterly.keyMetrics', '一頁摘要「驗證點」列出的指標（2–3 項；metrics 的 key）', '文字清單', C),
 ('quarterly.driver', '拆分依據：type＝mw 時依 MW 內插（endStart＝第 1 期期初 Accepted MW、endStartNote 來源）；type＝none（沒有 MW 資料的公司）時營收依 revenueSplit、CapEx 平分，MW 列顯示「不適用」', '物件', C),
 ('quarterly.revenueSplit', '各所屬模型期的營收拆法：anchor＝從 revenueAnchor 逐季線性爬升、driverAvg＝依平均在役 MW、equal＝平分', '文字清單', C),
 ('quarterly.revenueAnchor', '營收拆法 anchor 的起點（最新一季實際營收）：value、label、tag', '物件', M),
 ('quarterly.metrics', '追蹤的指標，一項一列：key（revenue、adjEbitda、ebitdaMargin、adjOpInc、capex、mw 之一）、label、unit（US$bn／%／MW）、gap 差距寫法（ratio＝比例，用於營收、CapEx；diff＝金額差；pt＝百分點）、marginOf 利潤類另列利潤率百分點時的分母指標（revenue）、consensus 共識檔季度欄位名（derived＝由共識 EBITDA ÷ 營收換算）、consensusSecondary／consensusSecondaryLabel 只列不比較的共識欄位、actualLabel 實際數列名稱', '清單', C),
 ('quarterly.capexSplit', 'CapEx 拆法：guidanceAnchor＝有季度指引的季取指引中點、其餘季分配期間餘數（依新增 MW）；driverAdds（或不填）＝全部依新增 MW', '文字', C),
 ('quarterly.guidanceText', '文字型指引（例如「調整後營業利益率 low teens」），{季別: {指標: {text, source, tag}}}：只列不計差距；完整支援（年增率、利潤率區間）列入待辦 5', '物件', M),
 ('quarterly.guidance', '季度指引（公司預估），{季別: {指標: [低, 高]}}；沒有指引的季度或公司不填，顯示「不適用」', '物件', M),
 ('quarterly.guidanceMeta', '季度指引的來源與標記', '物件', M),
 ('quarterly.periodGuidance', '期間隱含指引，{模型期: {指標: {range: [全年低, 全年高], less: 已實現部分}}}：只用於判斷季度差距是否為「拆法」', '物件', M),
 ('quarterly.periodGuidanceNote', '上一欄的說明文字', '文字', M),
 ('quarterly.consistency', '建置檢查：[路徑 A, 路徑 B, 倍數（選填）]，兩者（B × 倍數）必須相同，否則建置失敗（防止同一數字在兩處不一致）', '清單', C),
 ('quarterly.actuals', '季度實際數（公司公布後填入；預設空白＝待公布），{季別: {各指標, source, date, tag}}；Excel 對應「輸入與假設」J 區藍字格', '物件', M),
 ('varianceReasons._note', '差異原因的說明文字（不進程式）', '文字', K),
 ('varianceReasons.list', '差異原因（已決定事項 2），一筆一列：scope（annual 年度共識對照／quarter 季度）、period（FY27、2026Q3 或 *）、metric（年度：rev、ebitda、capex、nd；季度：metrics 的 key）、vs（consensus、guidance、actual 或 *）、type（觀點／已知限制）、text 一句原因，{路徑:格式} 由模型數字帶入。「拆法」由程式判定，不需填。差距超過 methodology.consensusGapTol 卻沒有原因時建置失敗', '清單', C),
 # Nebius v0.1a：公司專屬資料草稿（引擎尚未讀取；v0.1b 依各欄 mapTo 搬到既有欄位）
  ('oracle._readme', 'Oracle 資料草稿區段的說明（v0.1a 產出；引擎尚未讀取，v0.1b 依 mapTo 搬入既有欄位或新增）', '文字', M),
 ('oracle.files', 'Oracle 事實總帳、每 MW 推導、共識資料檔的路徑', '物件（路徑）', M),
 ('oracle.debt', 'Oracle 債務草稿（逐檔票券、定期貸款、商業本票、到期梯、循環額度、強制轉換特別股、信評）；格式同上', '物件', M),
 ('oracle.equity', 'Oracle ATM、FY27 融資計畫、股利、買回授權、股價草稿；格式同上', '物件', M),
 ('oracle.prepay', 'Oracle 客戶出資覆蓋比、預付累計、預付＋自帶硬體合約額草稿；格式同上', '物件', M),
 ('oracle.mw', 'Oracle 已交付 MW、站點、合約容量、利用率、續約溢價草稿（口徑逐欄註明）；格式同上', '物件', M),
 ('oracle.legacy', 'Oracle 傳統事業四線（SaaS、軟體、硬體、服務）近四季營收、成長率、分部利潤率草稿；格式同上', '物件', M),
 ('oracle.guidance', 'Oracle 公司指引草稿（Q2 FY27、FY27、OCI 路徑與 FY30 目標只作對照）；格式同上', '物件', M),
 ('oracle.valuation', 'Oracle 評價輸入草稿（beta、無風險利率、ERP、債務成本、稅率、同業倍數、淨負債、持股）；格式同上', '物件', M),
 ('oracle.perMw', 'Oracle 每 MW 年收入三情境（沿用 Nebius 的 Tokenomics 推導）、EBITDA 率、伺服器壽命草稿；格式同上', '物件', M),
 ('oracle.openai', 'Oracle–OpenAI 合約年額、隱含每 MW、OpenAI 計畫算力支出（只作對照）；格式同上', '物件', M),
 ('oracle.events', 'Oracle 評價日後事件（Project Jupiter 不可抗力通知）；格式同上', '物件', M),
 # MAG v0.1a：Microsoft 資料草稿（引擎尚未讀取；v0.1b′ 依各欄 mapTo 搬入 MAG 共用引擎）
 ('mag._readme', 'Microsoft 資料草稿區段的說明（MAG v0.1a 產出；引擎尚未讀取，v0.1b′ 依 mapTo 搬入 MAG 共用引擎欄位）', '文字', M),
 ('mag.files', 'Microsoft 事實總帳、k 證據、共識、Tokenomics 快照與名稱清單的路徑', '物件（路徑）', M),
 ('mag.calendar', '財年（6 月）、最新已申報季、期間標籤 FY27–FY31、評價錨定年草稿；每欄 value＋unit＋ref（總帳 id）＋mapTo', '物件', M),
 ('mag.segments', 'FY27 新分部（Agents and Infra、Devices and Consumer）近 8 季營收與營業利益、舊三分部年度對照、分部 D&A 說明草稿；格式同上', '物件', M),
 ('mag.lines', 'FY27 新產品別（Azure、M365 cloud、授權、Industry solutions、Frontier、搜尋與廣告、XBOX、Windows）近 8 季營收與近四季年增草稿；格式同上', '物件', M),
 ('mag.cloud', 'Azure 最新季年化、成長率、Microsoft Cloud 毛利率草稿；格式同上', '物件', M),
 ('mag.pl', 'FY24–FY26 營收、營業利益、D&A、SBC、利息、稅率、淨利、OpenAI 損益草稿；格式同上', '物件', M),
 ('mag.cashflow', 'FY24–FY26 營運現金流、現金資本支出、含融資租賃資本支出（季）、融資租賃、短期資產占比、FCF、回購、股利草稿；格式同上', '物件', M),
 ('mag.balance', 'FY26 末現金及短期投資、權益投資、PP&E、伺服器成本、應付資本支出草稿；格式同上', '物件', M),
 ('mag.debt', '債券面額、帳面、公允價值、到期梯、逐檔明細、商業本票草稿；格式同上', '物件', M),
 ('mag.leases', '營業／融資租賃負債、未起租 329.1B 與各季路徑、租賃成本、到期、改分類說明、建設與採購承諾草稿；格式同上', '物件', M),
 ('mag.rpo', '商用與全公司 RPO、12 個月比例、OpenAI 占比草稿；格式同上', '物件', M),
 ('mag.shares', '流通股、稀釋股、每股股利、回購授權與計畫內回購草稿；格式同上', '物件', M),
 ('mag.guidance', 'FY27 Q1 指引、FY27 全年方向、CY2026 資本支出、稅率、Azure 成長指引草稿（全部 [Interested-party]）；格式同上', '物件', M),
 ('mag.mw', '期初 AI 在役 MW 區間、對外占比、世代組合、neocloud 租用 MW、新增速度、公司目標、總容量對照草稿；格式同上', '物件', M),
 ('mag.kFactor', 'k_長約、k_現貨、長約占比、加權 k（低／基準／高）草稿；格式同上', '物件', M),
 ('mag.neocloud', '已揭露 neocloud 合約（Nebius、IREN、Nscale、Lambda、CoreWeave）與年租金草稿；格式同上', '物件', M),
 ('mag.related', 'OpenAI 持股、估值、營收、Azure 承諾、營收分成；Anthropic 承諾、投資、估值；OpenAI 模型 Azure 路徑（對照）草稿；格式同上', '物件', M),
 ('mag.split', 'AI／非 AI 雲端拆分試算（對外 AI 雲端收入、非 AI 殘差、上限檢查、影子收入）草稿；格式同上', '物件', M),
 ('mag.valuation', 'beta、無風險利率、債務成本、ERP、同業倍數中位數、AI 雲端倍數、現價草稿；格式同上', '物件', M),
 ('mag.consensus', '共識目標價、營收、資本支出摘要（全檔 data/consensus_msft_20261008.json）草稿；格式同上', '物件', M),
]
LQ = {
 'filed': ('申報日', '日期'), 'periodEnd': ('季末日', '日期'), 'revenue': ('當季營收', 'US$bn'), 'yoy': ('當季營收年增率', '比例'),
 'h1Revenue': ('上半年營收', 'US$bn'), 'costRev': ('當季營收成本', 'US$bn'), 'techInfra': ('當季技術與基礎設施費用', 'US$bn'),
 'opInc': ('當季 GAAP 營業損益', 'US$bn'), 'interest': ('當季利息費用', 'US$bn'), 'h1Interest': ('上半年利息費用', 'US$bn'),
 'ni': ('當季淨損益', 'US$bn'), 'h1Ni': ('上半年淨損益', 'US$bn'), 'epsDiluted': ('當季稀釋每股盈餘', 'US$'),
 'sbc': ('當季股份基礎薪酬', 'US$bn'), 'da': ('當季折舊攤銷', 'US$bn'), 'adjEbitda': ('當季調整後 EBITDA', 'US$bn'),
 'adjOpInc': ('當季調整後營業利益', 'US$bn'), 'rpo': ('季末 RPO', 'US$bn'), 'backlog': ('季末 backlog（含其他）', 'US$bn'),
 'rpo24m': ('RPO 於 24 個月內認列比例', '比例'), 'rpo25to48': ('RPO 於 25–48 個月認列比例', '比例'),
 'rpo49to78': ('RPO 於 49–78 個月認列比例', '比例'), 'cash': ('季末現金', 'US$bn'), 'restricted': ('受限現金', 'US$bn'),
 'marketable': ('有價證券', 'US$bn'), 'availability': ('未動用信用額度', 'US$bn'), 'debtPrincipal': ('債務本金合計', 'US$bn'),
 'recourseNet': ('追索債務淨額', 'US$bn'), 'nonRecourseNet': ('非追索債務淨額', 'US$bn'), 'ddtlOut': ('DDTL 未償餘額', 'US$bn'),
 'notesOut': ('票據與可轉債未償餘額', 'US$bn'), 'sharesA': ('A 股流通股數', 'bn 股'), 'sharesB': ('B 股流通股數', 'bn 股'),
 'sharesOut': ('流通股數合計', 'bn 股'), 'basicWaso': ('加權平均流通股數', 'bn 股'), 'capexQ2': ('當季資本支出', 'US$bn'),
 'capexH1': ('上半年資本支出', 'US$bn'), 'cashCapexH1': ('上半年現金購置固定資產', 'US$bn'), 'cfoH1': ('上半年營運現金流', 'US$bn'),
 'cashInterestH1': ('上半年現金利息', 'US$bn'), 'capInterestH1': ('上半年資本化利息', 'US$bn'), 'deferredTotal': ('遞延收入合計', 'US$bn'),
 'deferredIn': ('上半年遞延收入淨流入', 'US$bn'), 'ppe': ('固定資產毛額', 'US$bn'), 'cip': ('在建工程', 'US$bn'),
 'rouOp': ('營業租賃使用權資產', 'US$bn'), 'opLeaseLiab': ('營業租賃負債', 'US$bn'), 'finLeaseLiab': ('融資租賃負債', 'US$bn'),
 'onBalanceUndiscounted': ('已入帳租約未折現付款', 'US$bn'), 'offBalanceLease': ('未起租租約（表外）', 'US$bn'),
 'singleSiteCap': ('單一站點租金上限', 'US$bn'), 'singleSiteMw': ('該站點未交付 MW', 'MW'), 'constructionCostMw': ('按造價計租的未交付 MW', 'MW'),
 'equipCommitLo': ('設備採購承諾下緣', 'US$bn'), 'equipCommitHi': ('設備採購承諾上緣', 'US$bn'), 'jvCommit': ('JV 承諾出資上限', 'US$bn'),
 'jvPaidH1': ('上半年已付 JV 出資', 'US$bn'), 'vieExposure': ('可變利益實體（VIE）最大曝險', 'US$bn'),
 'rouObtainedH1': ('上半年新取得使用權資產', 'US$bn'), 'leaseCashH1': ('上半年租賃現金支付', 'US$bn'),
 'custA': ('最大客戶營收占比', '比例'), 'custB': ('第二大客戶營收占比', '比例'), 'custC': ('第三大客戶營收占比', '比例'),
 'debtIssuedH1': ('上半年借款', 'US$bn'), 'debtRepaidH1': ('上半年還款', 'US$bn'), 'equityH1': ('上半年股權募資', 'US$bn'),
}
CF = {
 'activeGw': ('主動電力', 'GW'), 'activeAddQ2': ('當季新增主動電力', 'MW'), 'juneAddMw': ('季末月單月新增', 'MW'),
 'contractedGw': ('簽約電力（法說日）', 'GW'), 'contractedGwQ2': ('簽約電力（季末）', 'GW'), 'poweredLandGw': ('已取得電力的土地／選擇權／意向書', 'GW'),
 'yeActiveGw': ('年底主動電力指引', 'GW'), 'gw2030': ('2030 年電力目標', 'GW'), 'dataCenters': ('資料中心數', '座'),
 'capexLo': ('全年資本支出指引下緣', 'US$bn'), 'capexHi': ('全年資本支出指引上緣', 'US$bn'),
 'nextQCapexLo': ('下一季資本支出指引下緣', 'US$bn'), 'nextQCapexHi': ('下一季資本支出指引上緣', 'US$bn'),
 'nextQIntLo': ('下一季利息指引下緣', 'US$bn'), 'nextQIntHi': ('下一季利息指引上緣', 'US$bn'),
 'revLo': ('全年營收指引下緣', 'US$bn'), 'revHi': ('全年營收指引上緣', 'US$bn'),
 'nextQRevLo': ('下一季營收指引下緣', 'US$bn'), 'nextQRevHi': ('下一季營收指引上緣', 'US$bn'),
 'adjOpLo': ('全年調整後營業利益指引下緣', 'US$bn'), 'adjOpHi': ('全年調整後營業利益指引上緣', 'US$bn'),
 'nextQAdjOpLo': ('下一季調整後營業利益指引下緣', 'US$bn'), 'nextQAdjOpHi': ('下一季調整後營業利益指引上緣', 'US$bn'),
 'arrLo': ('期末年化經常性收入（ARR）指引下緣', 'US$bn'), 'arrHi': ('期末 ARR 指引上緣', 'US$bn'),
 'postQNewCommit': ('季末後新增承諾', 'US$bn'), 'priceUp': ('新約漲價幅度', '比例'), 'marginStep': ('新約貢獻率提升', '文字'),
 'backlogStarted': ('backlog 中已開始交付的比例（法說「>50%」）', '比例'), 'inferenceArr': ('推論服務 ARR（年底）', 'US$bn'), 'otherArr': ('其他服務 ARR', 'US$bn'),
 'convert': ('期後可轉債發行額', 'US$bn'), 'convertNet': ('期後可轉債淨額', 'US$bn'), 'convertCoupon': ('期後可轉債票面利率', '比例'),
 'convertPx': ('轉換價', 'US$'), 'atmShares': ('股權分銷（ATM）上限', 'm 股'), 'price0917': ('9/17 收盤價', 'US$'),
 'tier34Active': ('主動電力位於 Tier 3/4 市場比例', '比例'), 'tier34Contracted': ('簽約電力位於 Tier 3/4 市場比例', '比例'),
 'cdsMid': ('CDS 中價', 'bps'), 'cdsBidAsk': ('CDS 買／賣價', '文字'), 'cdsRange': ('CDS 近期區間', '文字'), 'cdsPeak': ('CDS 峰值', 'bps'),
 'cdsPeakDate': ('CDS 峰值日期', '日期'), 'cdsLow': ('CDS 低點', 'bps'), 'cdsLowDate': ('CDS 低點日期', '日期'),
 'availability': ('未動用信用額度', 'US$bn'), 'price0918': ('9/18 收盤價', 'US$'), 'priceLast': ('最新收盤價', 'US$'),
 'priceDate': ('最新收盤價日期', '日期'), 'postQShortDated': ('期後公告摘要', '文字'), 'ttmRev': ('近十二個月營收', 'US$bn'),
 'ttmOpInc': ('近十二個月 GAAP 營業損益', 'US$bn'), 'ttmDa': ('近十二個月折舊攤銷', 'US$bn'),
 'ttmOpLease': ('近十二個月營業租賃成本', 'US$bn'), 'mktCapLast': ('本公司市值（Comps 用；與同業同一收盤日 peers.priceDate，v4.3 起不隨現價更新）', 'US$bn'), 'nextEarn': ('下次財報時間', '文字'),
}


def get(path):
    v = CO
    for k in path.split('.'): v = v[k]
    return v


def fmt(v):
    if isinstance(v, bool): return '是' if v else '否'
    if isinstance(v, list):
        if v and isinstance(v[0], (dict, list)): return f'{len(v)} 筆'
        return '、'.join(fmt(x) for x in v)
    if isinstance(v, dict): return '物件（' + '、'.join(v) + '）'
    if isinstance(v, str): return v if len(v) <= 28 else v[:26] + '…'
    return f'{v:g}' if isinstance(v, float) else str(v)


# 覆蓋檢查：company.json 每個欄位都要有說明
doc = {p for p, *_ in F} | {f'latestQuarter.{k}' for k in LQ} | {f'callFacts.{k}' for k in CF}
def leaves(o, p=''):
    if isinstance(o, dict):
        for k, v in o.items():
            q = f'{p}.{k}' if p else k
            if q in doc: yield q
            else: yield from leaves(v, q)
    else: yield p
missing = [q for q in leaves(CO) if q not in doc]
assert not missing, f'未說明的欄位：{missing}'
for q in doc: get(q)  # 說明了不存在的欄位會在此報錯

SECT = [('meta', '基本資料'), ('calendar', '期間與日期（v4.5）'), ('asOf', '滾動檢查（首期一次性金額與期初餘額的所屬季度）'), ('ytdActual', '年初至今實際數（10-Q；v4.5 前為 actual1H）'), ('historicalPL', '歷年損益'),
        ('rpo', '已簽約未認列營收（RPO）'), ('leases', '租約'), ('debt', '既有債務'), ('latestQuarter', '最新一季財報數字（10-Q）'),
        ('callFacts', '法說會與期後事項'), ('scenarios', '三個擴張情境'), ('legacy', '舊版對照值'),
        ('defaults', '預設假設（畫面上可調的輸入）'), ('valuation', '評價參數'), ('methodology', '評價方法與評等門檻'), ('peers', '同業比較（Comps）'),
        ('quarterly', '季度層（v4.4）'), ('varianceReasons', '差異原因（v4.4）'), ('texts', '公司特有的說明文字（v4.5；隨資料更新）'), ('oracle', 'Oracle 資料草稿（v0.1a；v0.1b 逐步搬入，只留後續步驟用或只作對照的欄位）'), ('mag', 'Microsoft 資料草稿（MAG v0.1a；引擎尚未讀取）')]
out, shown = ['**填表慣例**',
               '- 金額單位是**十億美元（US$bn）**，例如 4.653 代表 46.53 億美元；另有標示的例外：每股（US$）、每 MW 建置成本（百萬美元／MW，US$m/MW）、股數（十億股，bn）。',
               '- 「比例」寫成小數（0.25＝25%）；標示「%」的欄位寫成百分點（25＝25%）。兩種寫法沿用既有程式，不可混用。',
               f"- 「清單」依模型期順序填：{'、'.join(PL)}，共 {len(PL)} 格（除非另有說明）。",
               '- 文字中的來源標記沿用 [Verified]（已公開可查）、[Interested-party]（利害關係人說法）、[Derived]（由其他數字換算）、[Assumed]（判斷值）。',
               '- 「換公司」欄：**必改**＝公司特有的資料；**檢查**＝判斷值，要依新公司重新評估；**可沿用**＝口徑或方法，通常不必改。',
               f"- 下表的「目前數值」是 {CO['meta']['company']} {VER} 的值（版本號讀 `vlog.py`、期間讀 `calendar_q.py`，由本檔自動帶入）；過長的文字只顯示開頭。表格由 `scripts/fields_doc.py` 產生，新增欄位時先在該檔補說明，再重新產生。"], set()
for top, title in SECT:
    out.append(f'\n### `{top}`：{title}\n')
    if top in ('latestQuarter', 'callFacts'):
        acc = 'LATEST_Q.' if top == 'latestQuarter' else 'CALL_FACTS.'
        tbl = LQ if top == 'latestQuarter' else CF
        out.append('換公司：全部必改（公司特有資料）。「程式使用」為否的欄位只是存檔備查，畫面與 Excel 的說明文字目前另外寫在程式裡。\n')
        out.append('| 欄位 | 意義 | 單位 | 目前數值 | 程式使用 |\n|---|---|---|---|---|')
        for k, (m, u) in tbl.items():
            used = re.search(rf"({re.escape(acc)}|{top}'\]\['){k}\b", CODE) is not None
            out.append(f'| `{k}` | {m} | {u} | {fmt(CO[top][k])} | {"是" if used else "否"} |')
        continue
    out.append('| 欄位 | 意義 | 單位 | 目前數值 | 換公司 |\n|---|---|---|---|---|')
    for p, m, u, c in F:
        if p == top or p.startswith(top + '.'):
            shown.add(p)
            out.append(f'| `{p}` | {m} | {u} | {fmt(get(p))} | {c} |')
assert shown == {p for p, *_ in F}, set(p for p, *_ in F) - shown
txt = '\n'.join(out)
if sys.argv[1:2] == ['--check']:  # 5a：README「填表慣例」起至各區表格結束，須與本檔輸出相同（verify.sh 步驟 0b）
    R = open(os.path.join(REPO, 'README.md'), encoding='utf-8').read().split('\n')
    a = next(i for i, l in enumerate(R) if l.startswith('**填表慣例**'))
    b = next(i for i, l in enumerate(R) if i > a and l.startswith('## '))
    cur = '\n'.join(R[a:b]).strip('\n')
    if cur != txt.strip('\n'):
        import difflib
        print('README 欄位說明與 company.json／fields_doc.py 不一致（執行 python3 scripts/fields_doc.py --write 更新）：')
        print('\n'.join(list(difflib.unified_diff(cur.split('\n'), txt.strip('\n').split('\n'), lineterm='', n=0))[:20])); sys.exit(1)
    print('README 欄位說明與 company.json 一致'); sys.exit(0)
if sys.argv[1:2] == ['--write']:  # 直接更新 README 的對應區段
    R = open(os.path.join(REPO, 'README.md'), encoding='utf-8').read().split('\n')
    a = next(i for i, l in enumerate(R) if l.startswith('**填表慣例**'))
    b = next(i for i, l in enumerate(R) if i > a and l.startswith('## '))
    open(os.path.join(REPO, 'README.md'), 'w', encoding='utf-8').write('\n'.join(R[:a] + txt.strip('\n').split('\n') + [''] + R[b:]))
    print('README 欄位說明已更新'); sys.exit(0)
print(txt)
