# MiniMax v0.2.3（v0.2.5 換價）：估值頁 Val 與反向 DCF 頁 RevDCF（Andy 2026-11-15 決定 MiniMax 與智譜都給目標價與評等）。
# 排版與公式結構在 valcore.py（與 zhipu/builder/valcore.py 同一份程式）；本檔只提供參照與新增的 SRC／Inputs 列。
# - 現金流沿用模型既有路徑：營收 R09、EBIT＝PnL P16 營業虧損（含股份支付，不加回）、資本支出＝Inputs I22（算力預付與自建）；
#   2H26＝FY2026 − 1H26 實績（營收 S004；營業虧損＝營收 − 銷售成本 − 三項費用，S004、S020、S024–S026）。
#   算力為雲端租用（合約價口徑），已在 EBIT 扣除（折舊 − 資本支出＝0）。
# - 評價日 2026-06-30（中期報告）；其後完成的 2026-07 配售與可轉債淨額計入淨現金；股數取 2026-09-30 FF301（352.48M 股）；
#   現價 2026-11-20 收盤 HK$102.4（S113；v0.2.4 為 11-13 收盤 152.4，v0.2.3 為 11-05 收盤 170.2）。可轉債轉換價 HK$335 高於現價 → 視為債務。
# - 情境 A／C：API token 量年增（I09）取高／低端（與 Sensitivity 頁同兩案），快照 data/val_scenarios.json（scripts/val_snapshots.py）。
# - 同業：智譜（2513.HK），市值與 MiniMax 同日（2026-11-20 收盤；v0.2.4 為 11-13、v0.2.3 為 11-05）；其餘候選找不到同日可查證市值，未納入（README 列待確認）。
# 兩頁不被任何既有計算頁引用；新增後既有工作表數值不變（tools/xlsx_diff.py 逐格驗證）。
import json
import os

import valcore

HERE = os.path.dirname(os.path.abspath(__file__))
SNAP_JSON = os.path.join(HERE, 'data', 'val_scenarios.json')
SENS_JSON = os.path.join(HERE, 'data', 'val_sens.json')
ZP = '智譜 2513.HK'
ZPSRC = '智譜模型 zhipu/data/zhipu_src.yaml'

# (ID, 項目, 期間, 值, 單位, 標記, 一手/二手, 來源) — 附加在 SRC 尾端（既有列號不變）
SRC_ADD = [
    ('S142', '已發行股份（A 類 271,377,649＋B 類 81,102,534）', '2026-09-30', 352.480183, 'M 股', 'Verified', '一手',
     'MiniMax FF301 月報表 https://pdf.dfcfw.com/pdf/H2_AN202610021830089929_1.pdf（智譜模型 SRC_ZP_597 同一文件）'),
    ('S143', f'{ZP} 收盤價', '2026-11-20', 334.0, 'HK$/股', 'Verified', '二手',
     '工作單 px1120；Yahoo Finance 2513.HK 日線收盤 2026-11-20（markets/invest_map/data/closes_20261109_20261120.csv）：334.0；v0.2.4 為 2026-11-13 收盤 408.4；v0.2.3 為 2026-11-05 收盤 476.0（Yahoo Finance 日線；StockAnalysis HKG:2513 歷史頁同值，智譜模型 SRC_ZP_1055）'),
    ('S144', f'{ZP} 已發行股份總數（2026-09 配售後）', '2026-09-16', 487.58809, 'M 股', 'Verified', '一手', f'智譜 2026-09-13 配售公告（{ZPSRC} shares_total_post_placing2）'),
    ('S145', f'{ZP} 現金及現金等價物', '2026-06-30', 39.9372, 'RMB 億', 'Verified', '一手', f'智譜 2026 中期業績公告（{ZPSRC} SRC_ZP_211）'),
    ('S146', f'{ZP} 2026-07 H 股配售淨額', '2026-07-13', 31374.95, 'HK$M', 'Verified', '一手', f'智譜 2026-09-13 公告（{ZPSRC} SRC_ZP_536）'),
    ('S147', f'{ZP} 2026-09 H 股配售淨額', '2026-09-16', 15664.13, 'HK$M', 'Verified', '一手', f'智譜 2026-09-13 公告（{ZPSRC} SRC_ZP_544）'),
    ('S148', f'{ZP} 可換股債券淨額', '2026-09-18', 3010.77, '$M', 'Verified', '一手', f'智譜 2026-09-13 公告（{ZPSRC} SRC_ZP_547）'),
    ('S149', f'{ZP} 銀行借款', '2026-06-30', 22.2479, 'RMB 億', 'Verified', '一手', f'智譜 2026 中期業績公告（{ZPSRC} SRC_ZP_213）'),
    ('S150', f'{ZP} 可換股債券本金（換股價 HK$892.5，價外）', '2026-09-18', 201.4, 'RMB 億', 'Verified', '一手', f'智譜 2026-09-13 公告（{ZPSRC} SRC_ZP_545）'),
    ('S151', f'{ZP} 租賃負債（1H26 未揭露，沿用最近一期）', '2025-12-31', 5.4514, 'RMB 億', 'Verified', '一手', f'智譜 2025 年度業績公告（{ZPSRC} lease_liab_fy25）'),
    ('S152', f'{ZP} 共識營收（Simply Wall St，23 位）', 'FY2026E', 60.48, 'RMB 億', 'Interested-party', '二手', f'simplywall.st 2026-10-06（{ZPSRC} cons_rev_fy26）'),
    ('S153', f'{ZP} 共識營收（Simply Wall St，23 位）', 'FY2027E', 180.13, 'RMB 億', 'Interested-party', '二手', f'simplywall.st 2026-10-06（{ZPSRC} cons_rev_fy27）'),
    ('S154', f'{ZP} 模型匯率 HKD/CNY 中間價', '2026-09-28', 0.85923, 'RMB/HK$', 'Verified', '二手（含一手轉載）', f'CFETS 中間價（{ZPSRC} 的 Inputs fx_hkdcny）；同業倍數以智譜模型同一匯率換算'),
    ('S155', f'{ZP} 模型匯率 USD/CNY 中間價', '2026-09-28', 6.7399, 'RMB/$', 'Verified', '二手（含一手轉載）', f'CFETS 中間價（{ZPSRC} 的 Inputs fx_usdcny）'),
]

# 純量：(ID, 項目, 單位, 值, 低, 高, 標記, 說明) — 附加在 INPUTS 尾端
INP_ADD = [
    ('I35', '估值：無風險利率 rf（美國 10 年期公債 CMT）', '比例', 0.0486, 0.0386, 0.0586, 'Analogy',
     '2026-11-06 4.86%（ycharts 轉 Treasury CMT；與 Anthropic v0.2.0 SRC_ANT_576 同一來源）；港股港元計價、聯繫匯率，沿用美元 rf（判斷類，待 Andy 確認）'),
    ('I36', '估值：β（Blume 調整後）', '倍', 1.248, 1.0, 1.6, 'Analogy',
     'Anthropic v0.2.0 同業（PLTR、CRWD、NET、SNOW、DDOG）5 年 β 中位數 1.37 → Blume 1.248；MiniMax 2026-01 上市，自身 β 期間不足（判斷類）'),
    ('I37', '估值：股權風險溢酬 ERP', '比例', 0.05, 0.04, 0.06, 'Assumed', '個股DCF標準規格第 2 節預設 5%（4–6%）；未另加中國國家風險溢酬（判斷類，Sens 6% 涵蓋）'),
    ('I38', '估值：稅前債務成本利差（kd＝rf＋利差）', '比例', 0.02, 0.01, 0.04, 'Assumed', '樣板 Anthropic val_kd_spread；債務為銀行借款與零息可轉債，利差無揭露'),
    ('I39', '估值：稅率', '比例', 0.25, 0.15, 0.25, 'Assumed', '中國企業所得稅 25%（高新技術企業 15% 為低值）；扣評價日起自身累積虧損'),
    ('I40', '估值：WACC 調整（敏感度用）', '比例', 0.0, -0.01, 0.01, 'Decision', '基準 0；Sens ±1pt'),
    ('I41', '估值：永續成長率 g', '比例', 0.03, 0.02, 0.04, 'Assumed', '個股DCF標準規格第 3 節 g 3%（2–4%）'),
    ('I42', '估值：RONIC − WACC（0＝價值中性）', '比例', 0.0, 0.0, 0.05, 'Assumed', '個股DCF標準規格第 3 節 RONIC＝WACC'),
    ('I43', '估值：期中慣例與 2H26 半年（定義常數）', '年', 0.5, 0.5, 0.5, 'Decision', '定義'),
    ('I44', '估值：目標時點（評價日後年數）', '年', 1.35, 1.35, 1.35, 'Derived', '評價日 2026-06-30 至現價日 2026-11-05 約 0.35 年＋12 個月（v0.2.4、v0.2.5 換價時未重算，維持 1.35；實際約 1.37，工作單 px1113 只換價格）'),
    ('I45', '估值：延伸期 2033 年營收年增率（由 FY2030 年增線性收斂）', '比例', 0.10, 0.05, 0.15, 'Assumed', '樣板 Anthropic ext_g33 10%（5–15%）'),
    ('I46', '估值：延伸期營業利益率加碼（2033 年相對 FY2030，線性）', '比例', 0.0, 0.0, 1.1, 'Assumed',
     '樣板 Anthropic ext_m_add 基準 0（維持 FY2030 利益率）；高值 1.1＝2033 年利益率約 +13%（成熟軟體水準），只作敏感度'),
    ('I47', '估值：累積虧損扣抵無限開關（1＝不課稅）', '開關', 0, 0, 1, 'Decision', '樣板 Anthropic val_nol_inf_switch'),
    ('I48', '估值：營運資金率（占營收增量）', '比例', 0.0, -0.05, 0.05, 'Assumed', '樣板 Anthropic val_wc_rate 0（±5%）；1H26 合約負債增加（預收）方向為正向'),
    ('I49', '估值：NTM 營收中 FY2026 的權重（2026-11-05 起：2026 剩餘 56 天 ÷ 365；v0.2.4、v0.2.5 換價時未重算，工作單只換價格）', '比例', 0.153425, 0.153425, 0.153425, 'Derived', 'SpaceX v0.2.7 NTM 加權法；同業倍數與本模型 NTM 同一權重'),
    ('I50', '估值：Comparable 相對倍數調整開關（0＝不調整）', '開關', 0, 0, 1, 'Decision',
     '基準不調整：同業同為虧損中的 AI 模型公司（判斷類）；1＝倍數 × MIN(上限, 1＋FY2030 EBIT 利益率)，只作敏感度'),
    ('I51', '估值：Comparable 調整係數上限', '倍', 1.0, 1.0, 1.0, 'Decision', '樣板 Anthropic val_comp_cap'),
    ('I52', '估值：DCF 權重（其餘為 Comparable）', '比例', 0.5, 0.25, 0.75, 'Decision', '個股模型原則：DCF 與 Comparable 各半（判斷類）'),
    ('I53', '估值：情境 A 權重', '比例', 0.25, 0.25, 0.25, 'Decision', '個股DCF標準規格第 3 節 25／50／25'),
    ('I54', '估值：情境 B 權重', '比例', 0.5, 0.5, 0.5, 'Decision', '同上'),
    ('I55', '估值：情境 C 權重', '比例', 0.25, 0.25, 0.25, 'Decision', '同上'),
    ('I56', '評等邊界：強力買進（上檔空間 >）', '比例', 0.5, 0.5, 0.5, 'Decision', '原則第 24 條'),
    ('I57', '評等邊界：買進（上檔空間 >）', '比例', 0.25, 0.25, 0.25, 'Decision', '原則第 24 條'),
    ('I58', '評等邊界：中立（上檔空間 ≥）', '比例', -0.25, -0.25, -0.25, 'Decision', '原則第 24 條'),
    ('I59', '評等邊界：放空（上檔空間 ≥）', '比例', -0.5, -0.5, -0.5, 'Decision', '原則第 24 條'),
    ('I60', '評等：距邊界揭露門檻', '比例', 0.05, 0.05, 0.05, 'Decision', '原則第 36 條'),
    ('I61', '百分點換算（定義常數）', '倍', 100, 100, 100, 'Decision', '定義'),
]
KEYMAP = {'val_rf': 'I35', 'val_beta': 'I36', 'val_erp': 'I37', 'val_kd_spread': 'I38', 'val_tax': 'I39', 'val_wacc_adj': 'I40', 'val_g': 'I41',
          'val_ronic_add': 'I42', 'half': 'I43', 'val_t_tgt': 'I44', 'val_ext_g33': 'I45', 'val_ext_m_add': 'I46', 'val_nol_inf': 'I47',
          'val_wc_rate': 'I48', 'val_ntm_w26': 'I49', 'val_comp_adj_sw': 'I50', 'val_comp_cap': 'I51', 'val_w_dcf': 'I52', 'val_w_a': 'I53',
          'val_w_b': 'I54', 'val_w_c': 'I55', 'val_r_sb': 'I56', 'val_r_b': 'I57', 'val_r_n': 'I58', 'val_r_s': 'I59', 'val_r_near': 'I60', 'val_pct': 'I61'}


def _load(p):
    return json.load(open(p, encoding='utf-8')) if os.path.exists(p) else {}


def build(wb, REG, YCOL):
    S = lambda sid: f"SRC_MM!$E${REG[sid]['row']}"                       # noqa: E731
    I = lambda k: f"Inputs!$D${REG[KEYMAP.get(k, k)]['row']}"            # noqa: E731
    Y = lambda rid, y: f"{REG[rid]['sheet']}!${YCOL[y]}${REG[rid]['row']}"  # noqa: E731
    fx = I('I01')
    snap = _load(SNAP_JSON)
    sens = _load(SENS_JSON)
    op1h = f"({S('S004')}-{S('S020')}-{S('S024')}-{S('S025')}-{S('S026')})"
    src_B = {
        'rev': [f"={Y('R09', 2026)}-{S('S004')}"] + [f"={Y('R09', y)}" for y in range(2027, 2031)],
        'ebit': [f"={Y('P16', 2026)}-{op1h}"] + [f"={Y('P16', y)}" for y in range(2027, 2031)],
        'cx': [f"=MAX(0,-{Y('F07', 2026)}-({S('S050')}-{S('S049')}))"] + [f"=-{Y('F07', y)}" for y in range(2027, 2031)],
        'ma': ["=0"] * 5,
        'fy': [f"={Y('R09', y)}" for y in (2025, 2026, 2027)],
    }
    zfx, zufx = S('S154'), S('S155')
    cfg = dict(
        I=I, unit='$M', ps_unit='$／股', cur='美元',
        title='Val — 估值頁（v0.2.6；Andy 2026-11-15 決定給目標價與評等）：DCF（FCFF）50%＋同業 EV／NTM 營收 50%；情境 A／B／C 25／50／25；評價日 2026-06-30、現價 2026-11-20 收盤',
        sub=('FCFF＝EBIT（PnL P16 營業虧損，含股份支付不加回）− 稅（25% × 正 EBIT，扣評價日起自身虧損）− 資本支出（Inputs I22 算力預付與自建）− 營運資金增加。'
             '終值＝NOPAT(2033) ×(1 − g ÷ RONIC) ÷ (WACC − g)。情境 A、C 為快照值（藍字；scripts/val_snapshots.py），B 為活公式。判斷類參數（WACC 組成、同業、權重）依樣板慣例預設，待 Andy 確認。'),
        sub2='本頁不被任何既有計算頁引用。單位：百萬美元、百萬股、美元／股；目標價以港元表示（× 港幣兌美元 I01）。',
        val_date='2026-06-30（pro forma：計入其後完成的配售與可轉債）', px_date='2026-11-20',
        px=S('S113'), px_note='S113（Yahoo Finance 0100.HK 日線；StockAnalysis 同值）',
        pxfx=f"1/{fx}", pxfx_note='÷ 港幣兌美元 7.78（Inputs I01）',
        shares=S('S142'), shares_unit='M 股', shares_label='評價股數＝2026-09-30 已發行股份（FF301；百萬股）',
        shares_note='S142；選擇權與受限制股份未在模型登錄（資料缺口，未計稀釋）；可轉債價外，不計換股股數。既有 Valuation 頁 V02 仍用 S051（349.2M），不改',
        cash_items=[
            ('2026-06-30 現金結餘（公司口徑：現金＋定存＋理財）', S('S041'), 'S041'),
            ('＋2026-07 配售新股淨額', Y('F03', 2026), 'Funding F03'),
            ('＋2026-07 可轉債淨額', Y('F04', 2026), 'Funding F04'),
            ('− 計息銀行借款（2026-06-30）', f"-{S('S048')}", 'S048'),
            ('− 可轉債本金（價外，視為債務）', f"-{S('S044')}/{fx}", 'S044 ÷ I01'),
        ],
        debt=f"{S('S048')}+{S('S044')}/{fx}", debt_note='銀行借款＋可轉債本金；租賃負債未登錄（資料缺口）',
        rf_note='Inputs I35：美國 10 年期公債 4.86%（2026-11-06；同 Anthropic v0.2.0）',
        beta_note='Inputs I36：Anthropic v0.2.0 同業 β 中位數 Blume 調整 1.248（Analogy；判斷類）',
        ttgt_note='Inputs I44：評價日至現價日約 0.35 年＋12 個月',
        src_zh={'rev': '總營收（Revenue R09；2H26＝FY2026 − 1H26 實績）', 'ebit': 'EBIT＝營業虧損（PnL P16；2H26＝FY2026 − 1H26 實績）',
                'cx': '資本支出（Inputs I22 算力預付與自建；2H26 扣 1H26 已發生）', 'ma': '併購現金支出（無）'},
        scn_desc={'A': 'API token 量年增取高端（4.0／1.5／0.9／0.6）', 'B': '基準（活公式）', 'C': 'API token 量年增取低端（0.8／0.5／0.3／0.2）'},
        sens=sens,
        snap=snap.get('scenarios', {}), snap_note=f"{snap.get('model', '')}；scripts/val_snapshots.py",
        snap_check=(snap.get('scenarios', {}).get('B', {}) or {}).get('check'),
        src_B=src_B,
        ebit_note='PnL P16（含股份支付）；1H26 營業虧損＝營收 − 銷售成本 − 銷售及分銷 − 行政 − 研發（S004、S020、S024–S026）；理財收益不計（FCFF）',
        cx_note='Inputs I22：2026 年 322（1H26 物業廠房設備與預付款已增 322.3，2H26 取差額下限 0）；2027 起 0',
        ma_note='無已承諾併購',
        peer_date_note='市值一律 2026-11-20 收盤；淨現金為 2026-06-30 pro forma；共識營收為最近一次彙整',
        peers=[dict(name='智譜（2513.HK）', unit='RMB 億', px=S('S143'), pxfx=zfx, shares=f"{S('S144')}/100",
                    netcash=(f"{S('S145')}+{S('S146')}*{zfx}/100+{S('S147')}*{zfx}/100+{S('S148')}*{zufx}/100-{S('S149')}-{S('S150')}-{S('S151')}"),
                    c26=S('S152'), c27=S('S153'),
                    note='收盤價 2026-11-20 HK$334.0（v0.2.4 為 11-13 收盤 408.4；v0.2.3 為 11-05 收盤 476.0）；股數 2026-09 配售後；淨現金＝6-30 現金＋07／09 配售與可換股債券淨額 − 銀行借款 − 可換股債券本金 − 租賃負債（智譜模型同一匯率）；'
                         '共識 FY2026 60.48、FY2027 180.13 億元（Simply Wall St）；與智譜模型 v0.1.6 同一組數字')],
        mult_note='同業只有智譜一家（中國與海外上市 AI 模型／應用公司中，取得 2026-11-20 同日可查證市值者）；其餘候選未納入，見 README（判斷類，待 Andy 確認）',
        self=dict(c26=S('S062'), c27=S('S083'), note='stockanalysis 共識 FY2026 452（S062）、FY2027 1,030（S083）'),
        gate=f"SUM(Funding!$E${REG['F06']['row']}:$I${REG['F06']['row']})+SUM(Funding!$E${REG['F07']['row']}:$I${REG['F07']['row']})",
        gate_note='Funding F06 營運燒錢（經調整淨利）＋F07 算力預付與自建，FY2026–FY2030 合計',
        cons_tp=S('S084'), cons_label='Investing.com 23 位平均，2026-09-30',
        cons_note='S084（18 買、4 持有、1 賣）；日期早於現價日一個多月，僅作對照',
        method=('DCF（FCFF；EBIT＝PnL P16、股份支付不加回；稅 25% 扣自身虧損；WACC＝CAPM〔rf 4.86%、β 1.248、ERP 5%〕與債務加權；終值 RONIC＝WACC、g 3%；'
                '延伸期 2031–2033 營收年增收斂到 10%、利益率維持 FY2030）50%＋同業 EV／NTM 營收（智譜，同日市值；不做相對倍數調整）50%；情境 A／B／C 25／50／25'),
    )
    P = valcore.build_val(wb, cfg)
    R = valcore.build_revdcf(wb, cfg, P)
    return P, R
