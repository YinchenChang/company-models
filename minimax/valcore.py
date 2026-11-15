"""估值頁（Val）與反向 DCF 頁（RevDCF）的共用排版（智譜 v0.1.6、MiniMax v0.2.3；Andy 2026-11-15 決定兩家給目標價與評等）。

依據：docs/plan/個股DCF標準規格.md（FCFF）、docs/plan/個股模型原則_命題加營收分部.md（第 24、36、56、63 條）、
樣板 anthropic/builder/a11.py（Val）與 a12.py build_revdcf（RevDCF）。
本檔只產生結構與公式字串：所有數值由呼叫端給的儲存格參照（SRC／Inputs／既有計算頁）或建置快照（情境 A／C、敏感度）提供，
數值計算一律由 Excel（LibreOffice 重算）完成。兩家共用同一份程式（minimax/valcore.py 與 zhipu/builder/valcore.py 內容相同）。

欄：C＝2H26（評價日 2026-06-30 之後的半年）、D–G＝2027–2030（模型）、H–J＝2031–2033（延伸：營收年增由 FY2030 值線性收斂到 ext_g33；
EBIT 利益率＝FY2030 ＋ ext_m_add × 線性；資本支出占營收沿用 FY2030）。明確期現值到 2032；終值＝NOPAT(2033) ×(1 − g ÷ RONIC) ÷ (WACC − g)，以 2032 年底折現。
"""
from __future__ import annotations

from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

F_IN = Font(name="Arial", size=10, color="FF0000FF")
F_CALC = Font(name="Arial", size=10, color="FF000000")
F_LINK = Font(name="Arial", size=10, color="FF008000")
F_BOLD = Font(name="Arial", size=10, bold=True)
F_TITLE = Font(name="Arial", size=13, bold=True)
F_NOTE = Font(name="Arial", size=9, color="FF595959")
FILL_SEC = PatternFill("solid", fgColor="FFD9E1F2")
FILL_HDR = PatternFill("solid", fgColor="FFEDEDED")
SCN = ("A", "B", "C")
DC = "CDEFGHIJ"          # 2H26、2027–2033
SC = "CDEFG"             # 來源區：2H26、2027–2030
SRCK = ("rev", "ebit", "cx", "ma")
TIERS = ("強力買進", "買進", "中立", "放空", "強力放空")


def nm(wb, name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


def put(ws, ref, v, font=None, fmt=None, fill=None):
    c = ws[ref]
    c.value = v
    if font is None:
        font = (F_LINK if "!" in v else F_CALC) if isinstance(v, str) and v.startswith("=") else F_CALC
    c.font = font
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    return c


def header(ws, row, labels):
    for i, lab in enumerate(labels):
        c = ws.cell(row=row, column=1 + i, value=lab)
        c.font = F_BOLD
        c.fill = FILL_HDR


class Page:
    def __init__(self, wb, ws, prefix):
        self.wb, self.ws, self.pfx = wb, ws, prefix
        self.P = {}
        self.r = 5

    def row(self, key, label, unit, vals, cols=DC, note="", fmt="#,##0.000", font=None, name=None, span=None, bold=False):
        n = self.r
        put(self.ws, f"A{n}", label, F_BOLD if bold else F_CALC)
        put(self.ws, f"B{n}", unit, F_CALC)
        for j, v in enumerate(vals):
            if v is None:
                continue
            put(self.ws, f"{cols[j]}{n}", v, font, fmt=fmt)
        if note:
            put(self.ws, f"K{n}", note, F_NOTE)
        if key:
            self.P[key] = n
        if name:
            a, b = span or (cols[0], cols[0])
            nm(self.wb, name, f"{self.ws.title}!${a}${n}" + (f":${b}${n}" if b != a else ""))
        self.r += 1
        return n

    def sec(self, t):
        put(self.ws, f"A{self.r}", t, F_BOLD, fill=FILL_SEC)
        self.r += 1

    def c(self, key, col="C"):
        return f"${col}${self.P[key]}"


def median(rng, cnt):
    # pycel 不支援 MEDIAN：以 LARGE 取中位數（奇偶皆可）
    return f"(LARGE({rng},INT(({cnt}+1)/2))+LARGE({rng},INT({cnt}/2)+1))/2"


def build_val(wb, cfg):
    """cfg：見 zhipu/builder/z6.py、minimax/val_page.py。回傳 P（列號）。"""
    I = cfg["I"]                       # Inputs 鍵 → 儲存格參照
    U, PSU = cfg["unit"], cfg["ps_unit"]
    ws = wb.create_sheet("Val")
    put(ws, "A1", cfg["title"], F_TITLE)
    put(ws, "A2", cfg["sub"], F_NOTE)
    put(ws, "A3", cfg["sub2"], F_NOTE)
    pg = Page(wb, ws, "VAL")
    row, sec, c = pg.row, pg.sec, pg.c
    P = pg.P
    half = I("half")

    # ═══ 0. 結論（預留，最後回填） ═══
    sec("0. 結論（投資建議；四個目標價）")
    concl0 = pg.r
    pg.r += 30

    # ═══ 1. 現價、股數、淨現金 ═══
    sec(f"1. 現價、評價日股數與淨現金（評價日 {cfg['val_date']}；現價日 {cfg['px_date']}）")
    row("px", f"現價（{cfg['px_date']} 收盤；港元）", "HKD／股", [f"={cfg['px']}"], fmt="#,##0.00", name="VAL_Px", note=cfg["px_note"])
    row("pxfx", f"港元 → {cfg['cur']} 每股換算係數", "倍", [f"={cfg['pxfx']}"], fmt="0.00000", note=cfg["pxfx_note"])
    row("pxm", f"現價（{cfg['cur']}）＝港元現價 × 換算係數", PSU, [f"={c('px')}*{c('pxfx')}"], fmt="#,##0.00")
    row("shares", cfg["shares_label"], cfg["shares_unit"], [f"={cfg['shares']}"], fmt="#,##0.0000", name="VAL_Shares", note=cfg["shares_note"], bold=True)
    for k, (lab, f, note) in enumerate(cfg["cash_items"]):
        row(f"cash{k}", lab, U, [f"={f}"], note=note)
    nc0, nc1 = P["cash0"], P[f"cash{len(cfg['cash_items']) - 1}"]
    row("netcash", "淨現金（評價日 pro forma）＝上列合計", U, [f"=SUM(C{nc0}:C{nc1})"], name="VAL_NetCash", bold=True,
        note="完成於評價日後、現價日前的融資計入；債務（銀行借款、可轉債本金、租賃負債）扣除；可轉債換股價高於現價 → 視為債務")
    row("debt", "負債 D（WACC 權重用）＝銀行借款＋可轉債本金＋租賃負債", U, [f"={cfg['debt']}"], note=cfg["debt_note"])
    row("emkt", "股權市值 E＝現價 × 評價股數", U, [f"={c('pxm')}*{c('shares')}"], name="VAL_MktCap")
    row("ev_mkt", "市場 EV＝E − 淨現金", U, [f"={c('emkt')}-{c('netcash')}"], name="VAL_MktEV")

    # ═══ 2. 折現率 ═══
    sec("2. 折現率（CAPM；個股DCF標準規格第 2 節）")
    row("rf", "無風險利率 rf", "比例", [f"={I('val_rf')}"], fmt="0.00%", name="VAL_Rf", note=cfg["rf_note"])
    row("beta", "β（Blume 調整後）", "倍", [f"={I('val_beta')}"], fmt="0.000", name="VAL_Beta", note=cfg["beta_note"])
    row("erp", "股權風險溢酬 ERP", "比例", [f"={I('val_erp')}"], fmt="0.00%", name="VAL_ERP")
    row("ke", "股東權益成本 ke＝rf ＋ β × ERP", "比例", [f"={c('rf')}+{c('beta')}*{c('erp')}"], fmt="0.00%", name="VAL_Ke", bold=True)
    row("kd", "稅前債務成本 kd＝rf ＋ 利差", "比例", [f"={c('rf')}+{I('val_kd_spread')}"], fmt="0.00%")
    row("tax", "稅率", "比例", [f"={I('val_tax')}"], fmt="0.0%", name="VAL_Tax")
    row("kdat", "稅後債務成本＝kd ×（1 − 稅率）", "比例", [f"={c('kd')}*(1-{c('tax')})"], fmt="0.00%", name="VAL_KdAT")
    row("wacc", "WACC＝(E × ke ＋ D × 稅後 kd) ÷ (E＋D) ＋ WACC 調整", "比例",
        [f"=({c('emkt')}*{c('ke')}+{c('debt')}*{c('kdat')})/({c('emkt')}+{c('debt')})+{I('val_wacc_adj')}"], fmt="0.00%", name="VAL_WACC", bold=True)
    row("g", "永續成長率 g", "比例", [f"={I('val_g')}"], fmt="0.00%", name="VAL_g")
    row("ronic", "RONIC＝WACC ＋ 超額報酬差（0＝價值中性）", "比例", [f"={c('wacc')}+{I('val_ronic_add')}"], fmt="0.00%", name="VAL_RONIC")
    row("mid", "期中慣例（年中；定義常數）", "年", [f"={half}"], fmt="0.00")
    row("stub", "評價日至 2026 年底（年）", "年", [f"={half}"], fmt="0.00")
    row("ttgt", "目標時點（評價日後年數；以 ke 推進）", "年", [f"={I('val_t_tgt')}"], fmt="0.00", note=cfg["ttgt_note"])

    # ═══ 3. 情境來源 ═══
    sec("3. 情境來源（2H26、FY2027–FY2030；B＝活公式；A、C＝情境快照值〔藍字，建置快照〕）")
    header(ws, pg.r, ["項目", "單位", "2H26", "FY2027", "FY2028", "FY2029", "FY2030", "", "", "", "說明"])
    pg.r += 1
    zh = cfg["src_zh"]
    snaps = cfg.get("snap") or {}
    for s in SCN:
        sd = snaps.get(s) or {}
        put(ws, f"A{pg.r}", f"情境 {s}：{cfg['scn_desc'][s]}", F_BOLD)
        pg.r += 1
        live = s == "B" or not sd.get("values")
        for k in SRCK:
            vals = cfg["src_B"][k] if live else list(sd["values"][k])
            row(f"{s}_{k}", f"{zh[k]}（{s}）", U, vals, cols=SC, font=None if live else F_IN,
                note=("活公式" if s == "B" else (f"快照：{cfg['snap_note']}" if not live else "快照缺：暫用 B 活公式")))
        fy = cfg["src_B"]["fy"] if live else list(sd["values"]["fy"])
        row(f"{s}_fy", f"營收：FY2025、FY2026、FY2027 全年（{s}；NTM 與營運資金用）", U, fy, cols="CDE", font=None if live else F_IN)

    # ═══ 4. DCF ═══
    sec("4. DCF（FCFF；欄 C＝2H26，D–G＝2027–2030〔模型〕，H–J＝2031–2033〔延伸〕；明確期現值到 2032，2033 只供終值 NOPAT）")
    header(ws, pg.r, ["項目", "單位", "2H26", "2027", "2028", "2029", "2030", "2031", "2032", "2033", "說明"])
    pg.r += 1
    row("yn", "年度（數值；2H26 記 2026）", "年", list(range(2026, 2034)), fmt="0")
    row("t", "折現年數（評價日起；期中）", "年",
        [f"={c('stub')}*{c('mid')}"] + [f"={c('stub')}+({col}{P['yn']}-$D${P['yn']})+{c('mid')}" for col in DC[1:]], fmt="0.00")
    row("df", "折現因子 1 ÷ (1＋WACC)^年數", "—", [f"=1/(1+{c('wacc')})^{col}{P['t']}" for col in DC], fmt="0.0000")
    tx, nolinf = c("tax"), I("val_nol_inf")
    for s in SCN:
        put(ws, f"A{pg.r}", f"情境 {s}", F_BOLD, fill=FILL_SEC)
        pg.r += 1
        k = lambda key: f"d{s}_{key}"  # noqa: E731
        sv = lambda key, j: f"${SC[j]}${P[f'{s}_{key}']}"  # noqa: E731
        n_g = pg.r
        row(k("g"), "營收年增率（2027–2030 模型；2031–2033 由 FY2030 線性收斂到延伸期終點）", "比例",
            [None, f"={sv('rev', 1)}/$D${P[s + '_fy']}-1"] + [f"={sv('rev', j + 1)}/{sv('rev', j)}-1" for j in range(1, 4)] +
            [f"=$G{n_g}+({I('val_ext_g33')}-$G{n_g})*({col}{P['yn']}-$G${P['yn']})/($J${P['yn']}-$G${P['yn']})" for col in "HIJ"], fmt="0.0%")
        n_rev = pg.r
        row(k("rev"), "營收", U, [f"={sv('rev', j)}" for j in range(5)] + [f"={DC[i - 1]}{n_rev}*(1+{DC[i]}{n_g})" for i in range(5, 8)],
            name=f"VAL_Rev_{s}", span=("C", "J"))
        n_eb = pg.r + 1
        row(k("m"), "EBIT 利益率（2031–2033＝FY2030 ＋ 延伸期利益率加碼 × 線性）", "比例",
            [f"={col}{n_eb}/{col}{n_rev}" for col in "CDEFG"] +
            [f"=$G{pg.r}+{I('val_ext_m_add')}*({col}{P['yn']}-$G${P['yn']})/($J${P['yn']}-$G${P['yn']})" for col in "HIJ"], fmt="0.0%")
        row(k("ebit"), "EBIT（含股權報酬費用，不加回）", U, [f"={sv('ebit', j)}" for j in range(5)] + [f"={col}{n_rev}*{col}{n_eb - 1}" for col in "HIJ"],
            note=cfg["ebit_note"], name=f"VAL_EBIT_{s}", span=("C", "J"))
        n_no = pg.r
        row(k("nolo"), "累積虧損扣抵期初（評價日起自身虧損；之前虧損不計）", U, [f"=C{n_eb}-C{n_eb}"] + [f"={DC[i - 1]}{n_no + 2}" for i in range(1, 8)])
        row(k("tax"), "稅＝稅率 × MAX(0, EBIT − 扣抵期初) ×（1 − 扣抵無限開關）", U, [f"={tx}*MAX(0,{col}{n_eb}-{col}{n_no})*(1-{nolinf})" for col in DC])
        row(k("nolc"), "扣抵期末＝期初 ＋ 虧損 − 抵用", U, [f"={col}{n_no}+MAX(0,-{col}{n_eb})-MIN({col}{n_no},MAX(0,{col}{n_eb}))" for col in DC])
        row(k("nopat"), "NOPAT＝EBIT − 稅", U, [f"={col}{n_eb}-{col}{P[k('tax')]}" for col in DC])
        row(k("cx"), "− 資本支出（2031–2033 依 FY2030 占營收比例）", U,
            [f"={sv('cx', j)}" for j in range(5)] + [f"={col}{n_rev}*$G{pg.r}/$G{n_rev}" for col in "HIJ"], note=cfg["cx_note"])
        row(k("ma"), "− 已承諾併購現金支出（只在發生年）", U, [f"={sv('ma', j)}" for j in range(5)] + [f"={sv('ma', 0)}-{sv('ma', 0)}"] * 3, note=cfg["ma_note"])
        row(k("wc"), "− 營運資金增加＝營運資金率 × 營收增量（2H26＝(FY2026 − FY2025) × 半年）", U,
            [f"={I('val_wc_rate')}*($D${P[s + '_fy']}-$C${P[s + '_fy']})*{c('stub')}", f"={I('val_wc_rate')}*(D{n_rev}-$D${P[s + '_fy']})"] +
            [f"={I('val_wc_rate')}*({col}{n_rev}-{DC[i]}{n_rev})" for i, col in enumerate(DC[2:], start=1)])
        row(k("fcff"), "FCFF＝NOPAT − 資本支出 − 併購 − 營運資金增加", U,
            [f"={col}{P[k('nopat')]}-{col}{P[k('cx')]}-{col}{P[k('ma')]}-{col}{P[k('wc')]}" for col in DC], name=f"VAL_FCFF_{s}", span=("C", "J"), bold=True)
        row(k("pv"), "現值（到 2032）", U, [f"={col}{P[k('fcff')]}*{col}{P['df']}" for col in DC[:7]])

    # ═══ 5. 彙總 ═══
    sec("5. DCF 彙總與 Comparable（欄 C／D／E＝情境 A／B／C）")
    header(ws, pg.r, ["項目", "單位", "情境 A", "情境 B", "情境 C", "", "", "", "", "", "說明"])
    pg.r += 1
    X = dict(zip(SCN, "CDE"))
    s3 = lambda f: [f(s) for s in SCN]  # noqa: E731
    row("pvx", "明確期現值合計（2H26–2032）", U, s3(lambda s: f"=SUM(C{P['d' + s + '_pv']}:I{P['d' + s + '_pv']})"))
    row("tvn", "終值 NOPAT(2033)＝EBIT(2033) ×（1 − 稅率）", U, s3(lambda s: f"=J{P['d' + s + '_ebit']}*(1-{tx})"))
    row("tv", "終值（2032 年底）＝NOPAT(2033) ×(1 − g ÷ RONIC) ÷ (WACC − g)", U, s3(lambda s: f"={X[s]}{P['tvn']}*(1-{c('g')}/{c('ronic')})/({c('wacc')}-{c('g')})"))
    row("pvtv", "終值現值（折現年數＝2032 期中＋半年）", U, s3(lambda s: f"={X[s]}{P['tv']}/(1+{c('wacc')})^($I${P['t']}+{c('mid')})"))
    row("ev", "企業價值 EV", U, s3(lambda s: f"={X[s]}{P['pvx']}+{X[s]}{P['pvtv']}"), bold=True)
    row("tvsh", "終值占 EV（EV ≤ 0 時 N/M；原則第 56 條）", "比例", s3(lambda s: f'=IF({X[s]}{P["ev"]}>0,{X[s]}{P["pvtv"]}/{X[s]}{P["ev"]},"N/M")'), fmt="0.0%")
    row("evbad", "DCF 失效旗標：EV ≤ 0（1＝是）", "旗標", s3(lambda s: f"=IF({X[s]}{P['ev']}<=0,1,0)"), fmt="0")
    row("eqv", "股權價值＝EV ＋ 淨現金", U, s3(lambda s: f"={X[s]}{P['ev']}+{c('netcash')}"))
    row("ps0", "每股價值（評價日）＝股權價值 ÷ 評價股數", PSU, s3(lambda s: f"={X[s]}{P['eqv']}/{c('shares')}"), fmt="#,##0.00")
    row("ncps", "下限：淨現金每股（評價日）", PSU, s3(lambda s: f"={c('netcash')}/{c('shares')}"), fmt="#,##0.00", name="VAL_NetCashPS", span=("D", "D"),
        note="DCF 與 Comparable 統一下限（樣板 Anthropic v0.2.0）")
    row("dcfm", f"DCF 每股價值（目標時點；{cfg['cur']}）＝MAX(淨現金每股, 評價日) ×(1＋ke)^目標年數", PSU,
        s3(lambda s: f"=MAX({X[s]}{P['ncps']},{X[s]}{P['ps0']})*(1+{c('ke')})^{c('ttgt')}"), fmt="#,##0.00")
    row("dcf", "DCF 每股價值（目標時點；港元）＝上列 ÷ 換算係數", "HKD／股", s3(lambda s: f"={X[s]}{P['dcfm']}/{c('pxfx')}"), fmt="#,##0.00", bold=True)
    row("dcffl", "DCF 觸及淨現金下限（1＝是）", "旗標", s3(lambda s: f"=IF({X[s]}{P['ps0']}<{X[s]}{P['ncps']},1,0)"), fmt="0")
    for s in SCN:
        nm(wb, f"VAL_TP_DCF_{s}", f"Val!${X[s]}${P['dcf']}")
    nm(wb, "VAL_EV_B", f"Val!$D${P['ev']}")
    nm(wb, "VAL_PVExp_B", f"Val!$D${P['pvx']}")
    nm(wb, "VAL_PVTV_B", f"Val!$D${P['pvtv']}")
    nm(wb, "VAL_EqV_B", f"Val!$D${P['eqv']}")

    # ── 同業表 ──
    sec(f"5a. 同業表（{cfg['peer_date_note']}）：EV ÷ 未來 12 個月共識營收（NTM＝FY2026 剩餘天數 ÷ 365 × FY2026 ＋ 其餘 × FY2027）")
    header(ws, pg.r, ["同業", "單位", "現價（港元）", "每股換算係數", "股數", "淨現金", "共識營收 FY2026", "共識營收 FY2027", "EV", "EV／NTM 營收", "說明（出處、日期、幣別）"])
    pg.r += 1
    w26 = I("val_ntm_w26")
    p0 = pg.r
    for pr in cfg["peers"]:
        n = pg.r
        put(ws, f"A{n}", pr["name"], F_CALC)
        put(ws, f"B{n}", pr["unit"], F_CALC)
        for col, key, fmt in (("C", "px", "#,##0.00"), ("D", "pxfx", "0.00000"), ("E", "shares", "#,##0.0000"), ("F", "netcash", "#,##0.0"),
                              ("G", "c26", "#,##0.0"), ("H", "c27", "#,##0.0")):
            put(ws, f"{col}{n}", f"={pr[key]}", fmt=fmt)
        put(ws, f"I{n}", f"=C{n}*D{n}*E{n}-F{n}", fmt="#,##0.0")
        put(ws, f"J{n}", f"=I{n}/({w26}*G{n}+(1-{w26})*H{n})", fmt="0.00")
        put(ws, f"K{n}", pr["note"], F_NOTE)
        pg.r += 1
    p1 = pg.r - 1
    nm(wb, "VAL_PeerTbl", f"Val!$A${p0}:$J${p1}")
    rng = f"$J${p0}:$J${p1}"
    row("npeer", "同業家數", "家", [f"=COUNT({rng})"], fmt="0")
    row("mult", "同業 EV／NTM 營收中位數", "倍", [f"={median(rng, c('npeer'))}"], fmt="0.00", name="VAL_PeerMult", note=cfg["mult_note"])
    s = cfg["self"]
    row("selfntm", "對照：本公司共識 NTM 營收", U, [f"={w26}*{s['c26']}+(1-{w26})*{s['c27']}"], note=s["note"])
    row("selfmult", "對照：本公司市場 EV ÷ 共識 NTM 營收（現價隱含倍數）", "倍", [f"={c('ev_mkt')}/{c('selfntm')}"], fmt="0.00", name="VAL_SelfMult")
    row("relmult", "對照：本公司倍數 ÷ 同業中位數", "倍", [f"={c('selfmult')}/{c('mult')}"], fmt="0.00")

    # ── Comparable ──
    row("ntm", "本模型 NTM 營收（各情境）＝FY2026 剩餘權重 × FY2026 ＋ 其餘 × FY2027", U,
        s3(lambda s: f"={w26}*$D${P[s + '_fy']}+(1-{w26})*$E${P[s + '_fy']}"))
    nm(wb, "VAL_NTMRev", f"Val!$D${P['ntm']}")
    row("adj", "相對倍數調整係數（開關＝0 時 1；開關＝1 時＝MIN(上限, MAX(0, 1 ＋ FY2030 EBIT 利益率)))", "倍",
        s3(lambda s: f"=IF({I('val_comp_adj_sw')}=1,MIN({I('val_comp_cap')},MAX(0,1+$G${P['d' + s + '_m']})),1)"), fmt="0.000",
        note="基準不調整：同業同為虧損中的 AI 模型公司，倍數已反映同一獲利型態（判斷類，待 Andy 確認）")
    row("cmev", "Comparable EV＝同業倍數 × 調整係數 × 本模型 NTM 營收", U, s3(lambda s: f"={c('mult')}*{X[s]}{P['adj']}*{X[s]}{P['ntm']}"))
    row("cmpsm", f"Comparable 每股（{cfg['cur']}）＝MAX(淨現金每股, (EV ＋ 淨現金) ÷ 股數)", PSU,
        s3(lambda s: f"=MAX({X[s]}{P['ncps']},({X[s]}{P['cmev']}+{c('netcash')})/{c('shares')})"), fmt="#,##0.00")
    row("cmps", "Comparable 每股（港元）", "HKD／股", s3(lambda s: f"={X[s]}{P['cmpsm']}/{c('pxfx')}"), fmt="#,##0.00", bold=True)
    row("scen", "情境目標價＝DCF × 權重 ＋ Comparable ×（1 − 權重）", "HKD／股", s3(lambda s: f"={I('val_w_dcf')}*{X[s]}{P['dcf']}+(1-{I('val_w_dcf')})*{X[s]}{P['cmps']}"),
        fmt="#,##0.00")
    row("scenup", "情境上檔空間", "比例", s3(lambda s: f"={X[s]}{P['scen']}/{c('px')}-1"), fmt="0.0%")

    # ═══ 6. 收支關卡 ═══
    sec("6. 收支關卡（原則第 63 條；與 AI 投資地圖同一口徑）")
    gate = row("gate", "收支關卡：基準情境 2026–2030 自由現金流合計（負＝入不敷出）", U, [f"={cfg['gate']}"], name="VAL_Gate", note=cfg["gate_note"])
    row("gate2", "入不敷出判定用：收支合計", U, [f"=C{gate}"])
    row("gatef", "入不敷出判定（1＝不通過）", "旗標", [f"=IF(C{gate}<0,1,0)"], fmt="0", name="VAL_GateFail")

    # ═══ 7. 敏感度 ═══
    sec("7. 敏感度（龍捲風；一次改一個輸入，整本活頁簿 LibreOffice 重算後的加權目標價；建置快照）")
    header(ws, pg.r, ["案例", "低值情況下加權目標價", "高值情況下加權目標價", "低值設定", "高值設定", "低值：相對基準", "高值：相對基準", "白話名稱", "低值設定白話", "高值設定白話"])
    pg.r += 1
    s0 = pg.r
    sens = cfg.get("sens") or {}
    cases = sens.get("cases", [])
    sb_row = s0 + max(len(cases), 1)
    for cs in cases:
        n = pg.r
        put(ws, f"A{n}", cs["name"], F_CALC)
        put(ws, f"B{n}", cs.get("tp_lo"), F_IN, fmt="#,##0.00")
        put(ws, f"C{n}", cs.get("tp_hi"), F_IN, fmt="#,##0.00")
        put(ws, f"D{n}", cs.get("lo_label", ""), F_NOTE)
        put(ws, f"E{n}", cs.get("hi_label", ""), F_NOTE)
        put(ws, f"F{n}", f"=B{n}/$B${sb_row}-1", fmt="0.0%")
        put(ws, f"G{n}", f"=C{n}/$B${sb_row}-1", fmt="0.0%")
        put(ws, f"H{n}", cs.get("plain", cs["name"]), F_CALC)
        put(ws, f"I{n}", cs.get("lo_plain", cs.get("lo_label", "")), F_NOTE)
        put(ws, f"J{n}", cs.get("hi_plain", cs.get("hi_label", "")), F_NOTE)
        pg.r += 1
    if not cases:
        put(ws, f"A{pg.r}", "（尚未產生：執行敏感度快照工具）", F_NOTE)
        pg.r += 1
    nm(wb, "VAL_SensTbl", f"Val!$A${s0}:$J${max(s0, pg.r - 1)}")
    assert pg.r == sb_row, (pg.r, sb_row)
    row("sbase", "Sens 快照的基準加權目標價（建置時）", "HKD／股", [None], fmt="#,##0.00")
    put(ws, f"B{P['sbase']}", sens.get("base_tp"), F_IN, fmt="#,##0.00")
    put(ws, f"C{P['sbase']}", "=VAL_TP_W", fmt="#,##0.00")
    put(ws, f"K{P['sbase']}", f"B＝快照（{sens.get('model', '未產生')}）；C＝活公式；不同即過期", F_NOTE)
    row("sstale", "Sens 快照狀態（CURRENT／STALE）", "—", [f'=IF(ABS(B{P["sbase"]}-C{P["sbase"]})<0.000001,"CURRENT","STALE")'], fmt="@")
    P["sens_rows"] = (s0, sb_row - 1)
    if cfg.get("snap_check"):
        row("snapchk", "情境快照檢查：|活公式 B（營收＋EBIT 合計）− 快照 B|（應為 0）", U,
            [f"=ABS(SUM($C${P['B_rev']}:$G${P['B_rev']})+SUM($C${P['B_ebit']}:$G${P['B_ebit']})-{cfg['snap_check']!r})"], fmt="0.000000")
        row("snapst", "情境快照狀態（CURRENT／STALE）", "—", [f'=IF(C{P["snapchk"]}<0.000001,"CURRENT","STALE")'], fmt="@")
    end = pg.r

    # ═══ 0. 結論回填 ═══
    pg.r = concl0
    header(ws, pg.r, ["項目", "單位", "數值", "", "", "", "", "", "", "", "說明"])
    pg.r += 1
    W3 = lambda key: f"{I('val_w_a')}*$C${P[key]}+{I('val_w_b')}*$D${P[key]}+{I('val_w_c')}*$E${P[key]}"  # noqa: E731
    row("tpdcf", "DCF 目標價（三情境加權）", "HKD／股", [f"={W3('dcf')}"], fmt="#,##0.00", name="VAL_TP_DCF")
    row("tpcm", "Comparable 目標價（同業 EV／NTM 營收；三情境加權；基準未做相對倍數調整）", "HKD／股", [f"={W3('cmps')}"], fmt="#,##0.00", name="VAL_TP_Comp")
    row("tpw", "加權目標價＝DCF × 權重 ＋ Comparable ×（1 − 權重）", "HKD／股", [f"={I('val_w_dcf')}*{c('tpdcf')}+(1-{I('val_w_dcf')})*{c('tpcm')}"],
        fmt="#,##0.00", name="VAL_TP_W", bold=True)
    row("tpcons", f"市場共識目標價（{cfg['cons_label']}）", "HKD／股", [f"={cfg['cons_tp']}"], fmt="#,##0.00", name="VAL_TP_Cons", note=cfg["cons_note"])
    row("pxc", f"現價（{cfg['px_date']} 收盤）", "HKD／股", [f"={c('px')}"], fmt="#,##0.00")
    row("upw", "上檔空間：加權目標價 ÷ 現價 − 1", "比例", [f"={c('tpw')}/{c('px')}-1"], fmt="0.0%", name="VAL_Up_W", bold=True)
    row("updcf", "上檔空間：DCF", "比例", [f"={c('tpdcf')}/{c('px')}-1"], fmt="0.0%", name="VAL_Up_DCF")
    row("upcm", "上檔空間：Comparable", "比例", [f"={c('tpcm')}/{c('px')}-1"], fmt="0.0%", name="VAL_Up_Comp")
    row("upcons", "上檔空間：市場共識目標價", "比例", [f"={c('tpcons')}/{c('px')}-1"], fmt="0.0%", name="VAL_Up_Cons")
    u = c("upw")
    rate = (f'=IF({u}>{I("val_r_sb")},"強力買進",IF({u}>{I("val_r_b")},"買進",IF({u}>={I("val_r_n")},"中立",IF({u}>={I("val_r_s")},"放空","強力放空"))))')
    row("vrating", "估值評等（只看上檔空間）", "—", [rate], fmt="@", name="VAL_Rating", bold=True,
        note="邊界：>+50% 強力買進、>+25% 買進、≥−25% 中立、≥−50% 放空、其餘強力放空（原則第 24 條）")
    row("gview", "入不敷出判定", "—", [f'=IF(C{P["gate"]}<0,"不通過（收支合計 "&TEXT(C{P["gate"]},"0.0")&"）","通過")'], fmt="@")
    row("rec", "投資建議（收支關卡不通過且估值評等為買進以上時＝收支不足）", "—",
        [f'=IF(AND(C{P["gate"]}<0,OR({c("vrating")}="強力買進",{c("vrating")}="買進")),"收支不足",{c("vrating")})'], fmt="@", name="VAL_Rec", bold=True)
    row("dist", "距最近評等邊界（百分點）", "pt", [f"=MIN(ABS({u}-{I('val_r_sb')}),ABS({u}-{I('val_r_b')}),ABS({u}-{I('val_r_n')}),ABS({u}-{I('val_r_s')}))*{I('val_pct')}"], fmt="0.0",
        name="VAL_Dist")
    row("near", "距邊界 5pt 內（1＝是；原則第 36 條揭露）", "旗標", [f"=IF({c('dist')}<={I('val_r_near')}*{I('val_pct')},1,0)"], fmt="0")
    row("upflip", "升一級所需加權目標價（港元）", "HKD／股",
        [f"=IF({u}>{I('val_r_sb')},\"已最高\",{c('px')}*(1+IF({u}>{I('val_r_b')},{I('val_r_sb')},IF({u}>={I('val_r_n')},{I('val_r_b')},IF({u}>={I('val_r_s')},{I('val_r_n')},{I('val_r_s')})))))"],
        fmt="#,##0.00", name="VAL_UpFlip")
    row("wdcf_share", "DCF 在加權目標價的占比（金額）", "比例", [f"={I('val_w_dcf')}*{c('tpdcf')}/{c('tpw')}"], fmt="0.0%")
    row("dcf_floor_all", "三情境 DCF 皆觸及淨現金下限（1＝是）", "旗標", [f"=IF(SUM($C${P['dcffl']}:$E${P['dcffl']})=3,1,0)"], fmt="0", name="VAL_DCFFloorAll")
    note_f = ('="評等規則：加權目標價上檔空間 >+50% 強力買進、>+25% 買進、≥−25% 中立、≥−50% 放空、其餘強力放空。本次上檔空間 "&TEXT(' + u + '*100,"0.0")&"%（加權目標價 "&TEXT('
              + c("tpw") + ',"0.00")&"／現價 "&TEXT(' + c("px") + ',"0.00")&" 港元），估值評等"&' + c("vrating") + '&"；距最近邊界 "&TEXT(' + c("dist") + ',"0.0")&"pt"&IF('
              + c("near") + '=1,"（5pt 內：評等對假設敏感）","")&"；投資建議"&' + c("rec") + '&"。"')
    row("rnote", "評等說明（含距邊界揭露）", "—", [note_f], fmt="@", name="VAL_RatingNote")
    row("meth", "方法", "—", [cfg["method"]], fmt="@", name="VAL_Method")
    row("wsum", "檢查：情境權重合計 − 1（應為 0）", "—", [f"={I('val_w_a')}+{I('val_w_b')}+{I('val_w_c')}-1"], fmt="0.000000")
    row("tvmax", "檢查：三情境 DCF 失效（EV ≤ 0）個數", "個", [f"=SUM($C${P['evbad']}:$E${P['evbad']})"], fmt="0")
    assert pg.r <= concl0 + 30, pg.r
    ws.column_dimensions["A"].width = 72
    ws.column_dimensions["B"].width = 11
    for col in DC:
        ws.column_dimensions[col].width = 13
    ws.column_dimensions["K"].width = 80
    ws.freeze_panes = "C5"
    _ = end
    return P


def build_revdcf(wb, cfg, P):
    """RevDCF：現價隱含什麼（一次只動一個驅動）。(a) 支持現價所需 EV；(b) 終值期營業利益率（封閉解）；
    (c) 延伸期利益率加碼（封閉解近似：終值與 2031–2032 EBIT 皆隨加碼線性變動，以兩點差分求斜率）；(d) Comparable 隱含倍數與相對同業係數。
    WACC 與營收倍數的二分法反解不做：本模型基準 FCFF 各年為負，WACC 或營收規模變動無法使每股價值升到現價（無解，照實列出）。"""
    I = cfg["I"]
    ws = wb.create_sheet("RevDCF", index=wb.sheetnames.index("Val") + 1)
    put(ws, "A1", "RevDCF — 反向 DCF：現價隱含什麼（一次只動一個驅動、其餘維持基準；只讀 Val 頁與 Inputs，無任何頁引用本頁）", F_TITLE)
    put(ws, "A2", "DCF 以情境 B 路徑反解「情境 B 評價日每股價值 ×（1＋ke）^目標年數＝現價」。(b)(c)(d) 為封閉解；反解值超出合理範圍時照實列出，不外推。", F_NOTE)
    put(ws, "A3", f"單位：{cfg['unit']}、{cfg['ps_unit']}、港元／股。", F_NOTE)
    pg = Page(wb, ws, "RVD")
    R = pg.P
    V = lambda key, col="C": f"Val!${col}${P[key]}"  # noqa: E731

    def kv(key, label, f, fmt="#,##0.000", note="", name=None, bold=False):
        n = pg.r
        put(ws, f"A{n}", label, F_BOLD if bold else F_CALC)
        put(ws, f"B{n}", f, fmt=fmt)
        if note:
            put(ws, f"C{n}", note, F_NOTE)
        R[key] = n
        if name:
            nm(wb, name, f"RevDCF!$B${n}")
        pg.r += 1

    b = lambda key: f"$B${R[key]}"  # noqa: E731
    pg.sec("一、支持現價所需的 DCF 企業價值（情境 B 路徑）")
    kv("px", "現價（港元）", "=VAL_Px", "0.00", name="RVD_Px")
    kv("ps0r", f"支持現價所需的評價日每股價值（{cfg['cur']}）＝現價 × 換算係數 ÷（1＋ke）^目標年數", f"={V('pxm')}/(1+VAL_Ke)^{V('ttgt')}", "0.00", name="RVD_PS0Req")
    kv("ps0b", "模型：情境 B 評價日每股價值（未套下限）", f"={V('ps0', 'D')}", "0.00", name="RVD_PS0B")
    kv("evr", "支持現價所需的企業價值 EV＝每股 × 評價股數 − 淨現金", f"={b('ps0r')}*VAL_Shares-VAL_NetCash", name="RVD_EVReq", bold=True)
    kv("evb", "模型：企業價值 EV（情境 B）", "=VAL_EV_B", name="RVD_EVB")
    kv("evgap", "差距：所需 EV − 模型 EV", f"={b('evr')}-{b('evb')}", name="RVD_EVGap")
    pg.sec("二、現價隱含的終值期營業利益率（只改 2033 年營業利益率；終值 NOPAT 與其成正比；封閉解）")
    kv("mJ", "模型：2033 年營業利益率（情境 B；終值基礎）", f"=Val!$J${P['dB_m']}", "0.0%", name="RVD_TermMargin0")
    kv("m30", "對照：模型 2030 年營業利益率（情境 B）", f"=Val!$G${P['dB_m']}", "0.0%", name="RVD_Margin2030")
    kv("m1", "現價隱含的終值期營業利益率＝模型值 ×（終值現值 ＋ 所需 EV − 模型 EV）÷ 終值現值", f"={b('mJ')}*(VAL_PVTV_B+{b('evgap')})/VAL_PVTV_B", "0.0%",
       name="RVD_TermMargin", bold=True)
    kv("mchk", "核對：明確期現值 ＋ 終值現值 × 隱含利益率 ÷ 模型利益率 − 所需 EV（應為 0）", f"=VAL_PVExp_B+VAL_PVTV_B*{b('m1')}/{b('mJ')}-{b('evr')}", "0.000000", name="RVD_TermChk")
    pg.sec("三、現價隱含的延伸期利益率加碼（2031–2033 營業利益率由 FY2030 線性加碼到 2033；正值才有解）")
    # EV 對加碼為線性（加碼只進 2031–2033 EBIT 與終值；稅：2031 起若 EBIT 轉正需扣抵，近似以無扣抵計，列核對）
    kv("ext0", "模型：延伸期利益率加碼（Inputs）", f"={I('val_ext_m_add')}", "0.0%")
    kv("slope", "EV 對加碼的斜率（每 100% 加碼的 EV 變動；不計稅與扣抵：上限讀法）",
       "=(" + "+".join(f"Val!${col}${P['dB_rev']}*(Val!${col}${P['yn']}-Val!$G${P['yn']})/(Val!$J${P['yn']}-Val!$G${P['yn']})*Val!${col}${P['df']}" for col in "HI")
       + f"+Val!$J${P['dB_rev']}*(1-VAL_Tax)*(1-VAL_g/VAL_RONIC)/(VAL_WACC-VAL_g)/(1+VAL_WACC)^(Val!$I${P['t']}+Val!$C${P['mid']}))", "#,##0.0")
    kv("ext1", "現價隱含的延伸期利益率加碼（2033 年營業利益率 − FY2030；上限讀法）", f"={b('ext0')}+{b('evgap')}/{b('slope')}", "0.0%", name="RVD_ExtAdd", bold=True)
    kv("ext33", "對應的 2033 年營業利益率", f"={b('m30')}+{b('ext1')}", "0.0%", name="RVD_ExtMargin33")
    pg.sec("四、Comparable：現價隱含的 EV／NTM 營收倍數與相對同業係數（情境 B 的 NTM 營收；封閉解）")
    kv("cev", "現價隱含的企業價值（市場 EV）", "=VAL_MktEV", name="RVD_CompEV")
    kv("cm", "現價隱含的 EV ÷ 本模型 NTM 營收", f"={b('cev')}/VAL_NTMRev", "0.00", name="RVD_CompMult", bold=True)
    kv("pm", "模型：同業 EV／NTM 營收中位數", "=VAL_PeerMult", "0.00", name="RVD_PeerMult")
    kv("cc", "現價隱含的相對同業係數＝隱含倍數 ÷ 同業中位數", f"={b('cm')}/VAL_PeerMult", "0.000", name="RVD_CompCoef", bold=True)
    kv("ntmreq", "支持現價所需的本模型 NTM 營收（同業中位數倍數）", f"={b('cev')}/VAL_PeerMult", name="RVD_NTMReq")
    kv("ntmb", "模型：NTM 營收（情境 B）", "=VAL_NTMRev", name="RVD_NTMB")
    kv("cchk", "核對：（隱含倍數 × NTM 營收 ＋ 淨現金）÷ 股數 − 現價 × 換算係數（應為 0）", f"=({b('cm')}*VAL_NTMRev+VAL_NetCash)/VAL_Shares-{V('pxm')}", "0.000000", name="RVD_CompChk")
    pg.sec("五、不做的反解")
    put(ws, f"A{pg.r}", "WACC 與營收規模（等比例放大）的二分法反解不做：情境 B 的 2H26–2030 FCFF 皆為負、延伸期維持 FY2030 利益率，"
        "降低 WACC 或放大營收只會使 EV 更負，不存在使每股價值等於現價的解（Anthropic 樣板的 (a)(b) 不適用）。", F_NOTE)
    ws.column_dimensions["A"].width = 80
    ws.column_dimensions["B"].width = 16
    return R
