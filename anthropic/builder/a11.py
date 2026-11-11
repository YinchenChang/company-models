"""Anthropic v0.2.0 估值頁 Val（工作單 20261111 F1–F10；docs/plan/個股DCF標準規格.md FCFF 版；版面參考 spacex/build_spacex.py 的 Val 頁）。

M 線規格例外「AI 收支模型不加 DCF 腿」（2026-11-05）因公司掛牌而失效（J5 預設：DCF＋Comparable 各半）。
- 評價日 2026-09-30（pro forma 資產負債表日）；目標時點 12 個月（以 ke 推進）；現價＝ipo_px（IPO 區間中點 74）。
- FCFF＝EBIT（現金口徑＝Funding F06：營收淨額 − 算力現金 − 非算力〔不含股權報酬〕− 股權報酬；股權報酬不加回）− 稅 − 自建資本支出 − 營運資金增加。
  算力為服務／租賃付款，已在 EBIT 扣除 → 折舊 − 資本支出＝0（F03 自建資本支出另扣）。稅＝21% × 正 EBIT，扣除 2026Q4 起自身虧損累積的 NOL。
- 欄：2026Q4（＝FY2026 − 前九月實際代理）、2027–2030（模型）、2031–2033（延伸：營收淨額年增由 FY2030 值線性收斂到 ext_g33；EBIT 利益率維持 FY2030＋ext_m_add 線性）。
  明確期現值到 2032；終值＝NOPAT(2033) ×(1 − g ÷ RONIC) ÷ (WACC − g)，以 2032 年底折現。
- 情境 A／B／C（25／50／25）：B＝活公式；A、C＝情境快照值（tools/val_snapshots.py 以 LibreOffice 重算整本活頁簿後讀出，寫入 data/val_scenarios.yaml；
  只快照營運輸出〔營收、EBIT、資本支出、自由現金流率〕，折現率、股數、淨現金、倍數一律活公式）。Checks 比對快照 B 與活公式 B，過期立 WARN。
- Comparable：同業 EV／NTM 營收中位數 × NTM 營收總額（0.25 × FY2026＋0.75 × FY2027）；調整後＝倍數 × MIN(1, FY2030 自由現金流率 ÷ 同業 FCF 利益率中位數)。
- 敏感度表（龍捲風）：tools/val_sens.py 逐案改 Inputs、LibreOffice 重算，讀加權目標價寫入 data/val_sens.yaml（建置快照；Checks 比對基準值）。
中位數以 LARGE(範圍,(COUNT＋1)/2)（同業 5 家，奇數；pycel 不支援 MEDIAN）。數值一律引用 SRC_ANT／Inputs／既有具名範圍。
"""
from pathlib import Path

import yaml
from openpyxl.utils import get_column_letter as L

from common import F_BOLD, F_CALC, F_IN, F_NOTE, FILL_SEC, header, nm, put, title

ROOT = Path(__file__).resolve().parent.parent
SNAP_YAML = ROOT / "data" / "val_scenarios.yaml"
SENS_YAML = ROOT / "data" / "val_sens.yaml"
SCN = ("A", "B", "C")
SRCK = ("rev", "gross", "ebit", "own", "cc", "fcfm")           # 情境來源列（FY2026–FY2030）
SRC_ZH = {"rev": "營收淨額（截頂後；Funding F01）", "gross": "營收總額（截頂後；REV_GrossCapped）", "ebit": "EBIT（現金口徑）＝差額（Funding F06）",
          "own": "自建資本支出（Funding F03）", "cc": "算力成本（現金；Funding F02）", "fcfm": "自由現金流率（Funding F09）"}
PEERS = (("pltr", "PLTR", "Palantir"), ("crwd", "CRWD", "CrowdStrike"), ("net", "NET", "Cloudflare"), ("snow", "SNOW", "Snowflake"), ("ddog", "DDOG", "Datadog"))
PEER_DATE = {"ddog": "2026-10-30（頁面為 10-30 收盤）"}
DC = "CDEFGHIJ"                                                  # 2026Q4、2027–2033
SC = "CDEFG"                                                     # FY2026–FY2030（來源區）


def build(ctx):
    D, wb = ctx.D, ctx.wb
    S, I = D.S, D.I
    from a2 import Sheet
    F, Rv = Sheet.all["F"], Sheet.all["R"]
    snap = yaml.safe_load(SNAP_YAML.read_text(encoding="utf-8")) if SNAP_YAML.exists() else {}
    sens = yaml.safe_load(SENS_YAML.read_text(encoding="utf-8")) if SENS_YAML.exists() else {}
    ws = wb.create_sheet("Val")
    title(ws, "Val — 估值頁（v0.2.0；工作單 20261111 F）：DCF（FCFF）50%＋調整後 Comparable 50%；情境 A／B／C 25／50／25；評價日 2026-09-30、目標時點 12 個月；現價＝IPO 區間中點",
          "FCFF＝EBIT（現金口徑＝Funding F06，股權報酬不加回）− 稅（21% × 正 EBIT，扣 2026Q4 起自身 NOL）− 自建資本支出 − 營運資金增加；算力為服務／租賃付款已在 EBIT 扣除（折舊 − 資本支出＝0）。"
          "終值＝NOPAT(2033) ×(1 − g ÷ RONIC) ÷ (WACC − g)。情境 A、C 為快照值（藍字；tools/val_snapshots.py），B 為活公式。M 線規格例外「AI 收支模型不加 DCF 腿」因掛牌失效（J5）。",
          "本頁不被任何計算頁引用（Checks 除外）。單位：$B、十億股、$／股。")
    P = {}
    r = [5]

    def row(key, label, unit, vals, cols=DC, note="", fmt="#,##0.000", font=None, name=None, name_span=None, bold=False):
        n = r[0]
        put(ws, f"A{n}", label, F_BOLD if bold else F_CALC)
        put(ws, f"B{n}", unit, F_CALC)
        for j, v in enumerate(vals):
            if v is None:
                continue
            put(ws, f"{cols[j]}{n}", v, font, fmt=fmt)
        if note:
            put(ws, f"K{n}", note, F_NOTE)
        if key:
            P[key] = n
        if name:
            a, b = name_span or (cols[0], cols[0])
            nm(wb, name, f"Val!${a}${n}" + (f":${b}${n}" if b != a else ""))
        r[0] += 1
        return n

    def sec(t):
        put(ws, f"A{r[0]}", t, F_BOLD, fill=FILL_SEC)
        r[0] += 1

    def c(key, col="C"):
        return f"${col}${P[key]}"

    tax, stub_ = I("val_tax"), f"(1-{I('m_9m')}/{I('months_per_year')})"
    m2b = I("m_to_b")

    # ═══ 0. 結論 ═══
    sec("0. 結論（投資建議）")
    concl0 = r[0]
    r[0] += 22                                    # 預留結論區（最後回填）

    # ═══ 6. 同業表 ═══
    sec("1a. 同業表（工作單 F4、F6；stockanalysis.com；chat 端取得，建置代理照用）")
    header(ws, r[0], ["代號", "名稱", "EV／遠期營收", "TTM FCF 利益率", "5 年 β", "資料日"])
    r[0] += 1
    p0 = r[0]
    for kk, t, n_ in PEERS:
        n = r[0]
        put(ws, f"A{n}", t, F_CALC)
        put(ws, f"B{n}", n_, F_CALC)
        put(ws, f"C{n}", f"={S(f'peer_{kk}_evfs')}", fmt="0.00")
        put(ws, f"D{n}", f"={S(f'peer_{kk}_fcfm')}", fmt="0.0%")
        put(ws, f"E{n}", f"={S(f'peer_{kk}_beta')}", fmt="0.00")
        put(ws, f"F{n}", PEER_DATE.get(kk, "2026-11-06"), F_CALC)
        r[0] += 1
    nm(wb, "VAL_PeerTbl", f"Val!$A${p0}:$F${r[0] - 1}")
    P["peer0"], P["peer1"] = p0, r[0] - 1

    # ═══ 1. 折現率 ═══
    sec("1. 折現率（CAPM；債務只有融資租賃 1.95 → WACC≈ke）")
    row("rf", "無風險利率 rf（10 年美債 CMT，2026-11-06）", "比例", [f"={S('mkt_rf_cmt')}"], note="SRC_ANT_576（ycharts 轉 Treasury CMT）；第二來源 Tradeweb 11-09 4.872%（SRC_ANT_577）", fmt="0.00%",
        name="VAL_Rf")
    row("rf2", "對照：rf 第二來源（Tradeweb，2026-11-09）", "比例", [f"={S('mkt_rf_tw')}"], fmt="0.000%")
    pb = ",".join(S(f"peer_{k}_beta") for k, _, _ in PEERS)
    row("braw", "同業 5 年 β 中位數（PLTR、CRWD、NET、SNOW、DDOG）", "倍", [f"=LARGE($E${P['peer0']}:$E${P['peer1']},(COUNT($E${P['peer0']}:$E${P['peer1']})+1)/2)"], fmt="0.000", note="stockanalysis.com 統計頁 2026-11-06；單一來源")
    row("bblume", "Blume 調整 β＝w × 原始 ＋（1 − w）× 1", "倍", [f"={I('val_blume_w')}*{c('braw')}+(1-{I('val_blume_w')})*{I('val_blume_prior')}"], fmt="0.000")
    row("beta", "採用 β（val_beta_override＞0 時覆寫；Sens 1.0／1.6）", "倍", [f"=IF({I('val_beta_override')}>0,{I('val_beta_override')},{c('bblume')})"], fmt="0.000", name="VAL_Beta")
    row("erp", "股權風險溢酬 ERP", "比例", [f"={I('val_erp')}"], fmt="0.00%", name="VAL_ERP")
    row("ke", "股東權益成本 ke＝rf ＋ β × ERP", "比例", [f"={c('rf')}+{c('beta')}*{c('erp')}"], fmt="0.00%", name="VAL_Ke", bold=True)
    row("kd", "稅前債務成本 kd＝rf ＋ 利差（假設）", "比例", [f"={c('rf')}+{I('val_kd_spread')}"], fmt="0.00%")
    row("kdat", "稅後債務成本＝kd ×（1 − 稅率）", "比例", [f"={c('kd')}*(1-{tax})"], fmt="0.00%", name="VAL_KdAT")
    # 股數與淨現金先算（權重要用市值）
    sec("2. 評價日股數（十億股）與淨現金（工作單 F5）")
    px = I("ipo_px")
    row("px", "現價＝ipo_px（IPO 區間中點；上市後改收盤價）", "$/股", [f"={px}"], fmt="#,##0.00", name="VAL_Px")
    row("pxlo", "價格區間下緣", "$/股", [f"={S('pros_px_mid')}_Lo"], fmt="#,##0.00", name="VAL_PxLo")
    row("pxhi", "價格區間上緣", "$/股", [f"={S('pros_px_mid')}_Hi"], fmt="#,##0.00", name="VAL_PxHi")
    row("sh_post", "發行後股數（含首次公益發行、私募以 $74 計）", "十億股", [f"={S('pros_shares_post')}/{m2b}"], fmt="0.0000", note="SRC_ANT_539")
    row("sh_pp", "私募股數依價格調整＝26.0 ÷ 價格 − 351.35M", "十億股", [f"={S('pros_pp_total')}/{c('px')}-{S('pros_pp_shares_mid')}/{m2b}"], fmt="0.0000", note="SRC_ANT_525、533")
    row("sh_pbi", "扣：首次公益發行（val_pbi_switch＝0 時扣除；基準含）", "十億股", [f"=-(1-{I('val_pbi_switch')})*{S('pros_pbi_shares')}/{m2b}"], fmt="0.0000", note="SRC_ANT_541")
    row("sh_opt", "選擇權庫藏股法＝1,774.34M ×（1 − 2.91 ÷ 價格）", "十億股", [f"={S('pros_options')}/{m2b}*MAX(0,1-{S('pros_option_strike')}/{c('px')})"], fmt="0.0000",
        note="SRC_ANT_542、543")
    row("sh_rsu", "RSU（服務條件未滿 1,026.66M＋9-30 後新授 47.04M；全額）", "十億股", [f"=({S('pros_rsu_unvested')}+{S('pros_rsu_new')})/{m2b}"], fmt="0.0000", note="SRC_ANT_544、545")
    row("sh_vmo", "捐贈配對已歸屬義務 553.44M（val_vmo_switch；基準含）", "十億股", [f"={I('val_vmo_switch')}*{S('pros_vmo')}/{m2b}"], fmt="0.0000", note="SRC_ANT_546")
    row("sh_oa", "超額配售 123.75M（p_overallot_switch；基準 0）", "十億股", [f"={I('p_overallot_switch')}*{S('pros_overallot_shares')}/{m2b}"], fmt="0.0000", note="SRC_ANT_524")
    row("shares", "評價股數合計", "十億股", [f"=SUM(C{P['sh_post']}:C{P['sh_oa']})"], fmt="0.0000", name="VAL_Shares", bold=True)
    row("cash", "pro forma 現金與短投（9-30；依價格連動；已扣 RSU 代扣稅款）", "$B", [f"=FND_ProFormaCash"], note="Funding 第十六節 p_pfm；SRC_ANT_560")
    row("cash_oa", "＋超額配售淨額（p_overallot_switch；基準 0）", "$B", [f"={I('p_overallot_switch')}*Funding!$D${F.rows['p_oa']}"], note="Funding 第十六節 p_oa")
    row("debt", "− 融資租賃負債（表內）；循環額度已提領 0", "$B", [f"={S('pros_fin_lease')}"], note="SRC_ANT_566（工作單 F4）")
    row("netcash", "淨現金＝現金 ＋ 超額配售 − 融資租賃", "$B", [f"={c('cash')}+{c('cash_oa')}-{c('debt')}"], name="VAL_NetCash", bold=True)
    sec("1b. WACC")
    row("emkt", "股權市值 E＝現價 × 評價股數", "$B", [f"={c('px')}*{c('shares')}"], fmt="#,##0.0")
    row("dmkt", "負債 D＝融資租賃", "$B", [f"={c('debt')}"])
    row("wacc", "WACC＝(E × ke ＋ D × 稅後 kd) ÷ (E＋D) ＋ val_wacc_adj", "比例", [f"=({c('emkt')}*{c('ke')}+{c('dmkt')}*{c('kdat')})/({c('emkt')}+{c('dmkt')})+{I('val_wacc_adj')}"],
        fmt="0.00%", name="VAL_WACC", bold=True)
    row("g", "永續成長率 g", "比例", [f"={I('val_g')}"], fmt="0.00%", name="VAL_g")
    row("ronic", "RONIC＝WACC ＋ val_ronic_add（0＝價值中性）", "比例", [f"={c('wacc')}+{I('val_ronic_add')}"], fmt="0.00%", name="VAL_RONIC")
    row("mid", "期中慣例（年中）", "年", [f"={I('months_per_year')}/{I('months_per_year')}/{I('arith_k2')}"], fmt="0.00")
    row("stub", "評價日至 2026 年底（年）＝1 − 9 ÷ 12", "年", [f"={stub_}"], fmt="0.00")

    # ═══ 3. 情境來源（FY2026–2030） ═══
    sec("3. 情境來源（FY2026–FY2030；B＝活公式；A、C＝情境快照值〔藍字，tools/val_snapshots.py 以 LibreOffice 重算後讀出〕）")
    hr = r[0]
    header(ws, hr, ["項目", "單位", "FY2026", "FY2027", "FY2028", "FY2029", "FY2030", "", "", "", "說明"])
    r[0] += 1
    frow = {"rev": F.rows["rev"], "ebit": F.rows["gap"], "own": F.rows["own"], "cc": F.rows["cc"], "fcfm": F.rows["fcf_m"]}
    snaps = (snap or {}).get("scenarios", {})
    for s in SCN:
        sd = snaps.get(s, {})
        put(ws, f"A{r[0]}", f"情境 {s}：{sd.get('desc', '基準（活公式）' if s == 'B' else '快照缺：暫用 B 活公式')}", F_BOLD)
        r[0] += 1
        for k in SRCK:
            if s == "B" or not sd.get("values"):
                if k == "gross":
                    vals = [f"=Revenue!{'EFGHI'[j]}{Rv.rows['gross_c']}" for j in range(5)]
                else:
                    vals = [f"=Funding!{'EFGHI'[j]}{frow[k]}" for j in range(5)]
                font = None
            else:
                vals = list(sd["values"][k])
                font = F_IN
            row(f"{s}_{k}", f"{SRC_ZH[k]}（{s}）", "比例" if k == "fcfm" else "$B", vals, cols=SC, fmt="0.000" if k != "fcfm" else "0.0%", font=font,
                note=("活公式" if s == "B" else f"快照：{sd.get('id', '')}；{sd.get('model', '')}"))
    row("rev25", "FY2025 營收淨額（實際校準；三情境相同）", "$B", [f"=Funding!$D${F.rows['rev']}"])

    # ═══ 4. DCF ═══
    sec("4. DCF（FCFF；三情境各一區塊；欄 C＝2026Q4，D–G＝2027–2030〔模型〕，H–J＝2031–2033〔延伸〕；明確期現值到 2032，2033 只供終值 NOPAT）")
    hr = r[0]
    header(ws, hr, ["項目", "單位", "2026Q4", "2027", "2028", "2029", "2030", "2031", "2032", "2033", "說明"])
    r[0] += 1
    yrs = [I("year_2026")] + [I(f"year_{y}") for y in range(2027, 2034)]
    row("yr", "年度標籤", "—", ["2026Q4"] + [f"={y}" for y in yrs[1:]], fmt="0", name="VAL_Yr", name_span=("C", "J"))
    row("yn", "年度（數值）", "年", [f"={y}" for y in yrs], fmt="0")
    row("t", "折現年數（評價日起；期中）", "年", [f"={c('stub')}/{I('arith_k2')}"] + [f"={c('stub')}+({col}{P['yn']}-{I('year_2027')})+{c('mid')}" for col in DC[1:]], fmt="0.00")
    row("df", "折現因子 1 ÷ (1＋WACC)^年數", "—", [f"=1/(1+{c('wacc')})^{col}{P['t']}" for col in DC], fmt="0.0000", name="VAL_DF_B", name_span=("C", "J"))
    q4p = I("val_q4_proxy_switch")
    proxy = f"({q4p}*{S('pros_oi_9m26')}+(1-{q4p})*{S('pros_adjoi_9m26')})"
    BLK = {}
    for s in SCN:
        put(ws, f"A{r[0]}", f"情境 {s}", F_BOLD, fill=FILL_SEC)
        r[0] += 1
        k = lambda key: f"d{s}_{key}"  # noqa: E731
        sv = lambda key, j: f"${SC[j]}${P[f'{s}_{key}']}"  # noqa: E731
        # 營收淨額與成長
        row(k("g"), "營收淨額年增率（2027–2030 模型；2031–2033 由 FY2030 線性收斂到 ext_g33）", "比例",
            [None] + [f"={sv('rev', j + 1)}/{sv('rev', j)}-1" for j in range(4)] +
            [f"=$G{r[0]}+({I('ext_g33')}-$G{r[0]})*({col}{P['yn']}-{I('year_2030')})/({I('year_2033')}-{I('year_2030')})" for col in "HIJ"], fmt="0.0%")
        n_rev = r[0]
        row(k("rev"), "營收淨額", "$B",
            [f"={sv('rev', 0)}-{S('pros_rev_9m26')}*{sv('rev', 0)}/{sv('gross', 0)}"] + [f"={sv('rev', j)}" for j in range(1, 5)] +
            [f"={DC[i - 1]}{n_rev}*(1+{DC[i]}{P[k('g')]})" for i in range(5, 8)],
            note="2026Q4＝FY2026 − 前九月營收總額 33.242 × FY2026 淨額率（SRC_ANT_547）", name="VAL_Rev_B" if s == "B" else None, name_span=("C", "J"))
        n_eb = r[0] + 1
        row(k("m"), "EBIT 利益率（2031–2033＝FY2030 ＋ ext_m_add × 線性）", "比例",
            [f"=C{n_eb}/C{n_rev}"] + [f"={col}{n_eb}/{col}{n_rev}" for col in "DEFG"] +
            [f"=$G{r[0]}+{I('ext_m_add')}*({col}{P['yn']}-{I('year_2030')})/({I('year_2033')}-{I('year_2030')})" for col in "HIJ"], fmt="0.0%")
        row(k("ebit"), "EBIT（現金口徑；股權報酬不加回）", "$B",
            [f"={sv('ebit', 0)}-{proxy}"] + [f"={sv('ebit', j)}" for j in range(1, 5)] + [f"={col}{n_rev}*{col}{P[k('m')]}" for col in "HIJ"],
            note="2026Q4＝FY2026 F06 − 前九月代理（val_q4_proxy_switch＝1：GAAP 營業損益 −1.260〔SRC_ANT_548〕；0：調整後 4.081〔SRC_ANT_550〕）",
            name="VAL_EBIT_B" if s == "B" else None, name_span=("C", "J"))
        n_no = r[0]
        row(k("nolo"), "NOL 期初（2026Q4 起自身虧損累積；之前 NOL 不計）", "$B", [f"=C{n_eb}-C{n_eb}"] + [f"={DC[i - 1]}{n_no + 2}" for i in range(1, 8)])
        row(k("tax"), "稅＝稅率 × MAX(0, EBIT − NOL 期初) ×（1 − NOL 無限開關）", "$B",
            [f"={tax}*MAX(0,{col}{n_eb}-{col}{n_no})*(1-{I('val_nol_inf_switch')})" for col in DC], name="VAL_Tax_B" if s == "B" else None, name_span=("C", "J"))
        row(k("nolc"), "NOL 期末＝期初 ＋ 虧損 − 抵用", "$B", [f"={col}{n_no}+MAX(0,-{col}{n_eb})-MIN({col}{n_no},MAX(0,{col}{n_eb}))" for col in DC])
        row(k("nopat"), "NOPAT＝EBIT − 稅", "$B", [f"={col}{n_eb}-{col}{P[k('tax')]}" for col in DC])
        row(k("cx"), "− 自建資本支出（F03；2031–2033 依 FY2030 占營收比例）", "$B",
            [f"={sv('own', 0)}*{c('stub')}"] + [f"={sv('own', j)}" for j in range(1, 5)] + [f"={col}{n_rev}*{sv('own', 4)}/{sv('rev', 4)}" for col in "HIJ"],
            note="算力為服務／租賃付款，已在 EBIT 扣除；折舊 0")
        row(k("wc"), "− 營運資金增加＝val_wc_rate × 營收淨額增量（2026Q4＝(FY2026 − FY2025) × 0.25）", "$B",
            [f"={I('val_wc_rate')}*({sv('rev', 0)}-{c('rev25')})*{c('stub')}", f"={I('val_wc_rate')}*({sv('rev', 1)}-{sv('rev', 0)})"] +
            [f"={I('val_wc_rate')}*({col}{n_rev}-{DC[i]}{n_rev})" for i, col in enumerate(DC[2:], start=1)])
        row(k("fcff"), "FCFF＝NOPAT − 資本支出 − 營運資金增加", "$B", [f"={col}{P[k('nopat')]}-{col}{P[k('cx')]}-{col}{P[k('wc')]}" for col in DC],
            name="VAL_FCFF_B" if s == "B" else None, name_span=("C", "J"), bold=True)
        row(k("pv"), "現值（到 2032）", "$B", [f"={col}{P[k('fcff')]}*{col}{P['df']}" for col in DC[:7]])
        BLK[s] = k

    # ═══ 5. 彙總 ═══
    sec("5. DCF 彙總與 Comparable（欄 C／D／E＝情境 A／B／C）")
    header(ws, r[0], ["項目", "單位", "情境 A", "情境 B", "情境 C", "", "", "", "", "", "說明"])
    r[0] += 1
    X = dict(zip(SCN, "CDE"))
    s3 = lambda f: [f(s) for s in SCN]  # noqa: E731
    row("pvx", "明確期現值合計（2026Q4–2032）", "$B", s3(lambda s: f"=SUM(C{P['d' + s + '_pv']}:I{P['d' + s + '_pv']})"))
    row("tvn", "終值 NOPAT(2033)＝EBIT(2033) ×（1 − 稅率）（不抵 NOL）", "$B", s3(lambda s: f"=J{P['d' + s + '_ebit']}*(1-{tax})"))
    row("tv", "終值（2032 年底）＝NOPAT(2033) ×(1 − g ÷ RONIC) ÷ (WACC − g)", "$B", s3(lambda s: f"={X[s]}{P['tvn']}*(1-{c('g')}/{c('ronic')})/({c('wacc')}-{c('g')})"))
    row("pvtv", "終值現值（折現年數＝2032 期中＋0.5）", "$B", s3(lambda s: f"={X[s]}{P['tv']}/(1+{c('wacc')})^($I${P['t']}+{c('mid')})"))
    row("ev", "企業價值 EV", "$B", s3(lambda s: f"={X[s]}{P['pvx']}+{X[s]}{P['pvtv']}"), bold=True)
    row("tvsh", "終值占 EV（>80% 標示）", "比例", s3(lambda s: f"={X[s]}{P['pvtv']}/{X[s]}{P['ev']}"), fmt="0.0%")
    row("tvflag", "終值占 EV > 80%（1＝是）", "旗標", s3(lambda s: f"=IF({X[s]}{P['tvsh']}>0.8,1,0)"), fmt="0")
    row("eqv", "股權價值＝EV ＋ 淨現金", "$B", s3(lambda s: f"={X[s]}{P['ev']}+{c('netcash')}"))
    row("ps0", "每股價值（評價日）＝股權價值 ÷ 評價股數", "$/股", s3(lambda s: f"={X[s]}{P['eqv']}/{c('shares')}"), fmt="#,##0.00")
    row("dcf", "DCF 每股價值（目標時點）＝評價日 ×(1＋ke)^t", "$/股", s3(lambda s: f"=MAX(0,{X[s]}{P['ps0']})*(1+{c('ke')})^{I('val_t_tgt')}"), fmt="#,##0.00", bold=True)
    nm(wb, "VAL_EV_B", f"Val!$D${P['ev']}")
    nm(wb, "VAL_PVExp_B", f"Val!$D${P['pvx']}")
    nm(wb, "VAL_PVTV_B", f"Val!$D${P['pvtv']}")
    nm(wb, "VAL_TVShare_B", f"Val!$D${P['tvsh']}")
    nm(wb, "VAL_EqV_B", f"Val!$D${P['eqv']}")
    for s in SCN:
        nm(wb, f"VAL_TP_DCF_{s}", f"Val!${X[s]}${P['dcf']}")
    # Comparable
    row("ntm", "NTM 營收總額＝0.25 × FY2026 ＋ 0.75 × FY2027（評價日 2026-09-30）", "$B",
        s3(lambda s: f"={c('stub')}*$C${P[s + '_gross']}+(1-{c('stub')})*$D${P[s + '_gross']}"))
    nm(wb, "VAL_NTMRev", f"Val!$D${P['ntm']}")
    pe = ",".join(S(f"peer_{k}_evfs") for k, _, _ in PEERS)
    pf = ",".join(S(f"peer_{k}_fcfm") for k, _, _ in PEERS)
    row("mult", "同業 EV／遠期營收中位數（PLTR、CRWD、NET、SNOW、DDOG）", "倍", [f"=LARGE($C${P['peer0']}:$C${P['peer1']},(COUNT($C${P['peer0']}:$C${P['peer1']})+1)/2)"], fmt="0.00", name="VAL_PeerMult",
        note="stockanalysis.com，2026-11-06 收盤（DDOG 10-30）；EV／遠期營收＝遠期 P/S × EV／市值（chat 端推得）")
    row("pfcf", "同業 TTM FCF 利益率中位數", "比例", [f"=LARGE($D${P['peer0']}:$D${P['peer1']},(COUNT($D${P['peer0']}:$D${P['peer1']})+1)/2)"], fmt="0.0%", name="VAL_PeerFCFm")
    row("cmev", "市場主流 Comparable EV＝倍數 × NTM 營收總額（未調整）", "$B", s3(lambda s: f"={c('mult')}*{X[s]}{P['ntm']}"))
    row("cmps", "市場主流 Comparable 每股＝(EV ＋ 淨現金) ÷ 股數", "$/股", s3(lambda s: f"=MAX(0,({X[s]}{P['cmev']}+{c('netcash')})/{c('shares')})"), fmt="#,##0.00")
    row("adj", "調整係數＝MIN(上限, MAX(0, FY2030 自由現金流率 ÷ 同業 FCF 利益率中位數))（val_comp_adj_switch＝0 時 1）", "倍",
        s3(lambda s: f"=IF({I('val_comp_adj_switch')}=1,MIN({I('val_comp_cap')},MAX(0,$G${P[s + '_fcfm']}/{c('pfcf')})),1)"), fmt="0.000")
    nm(wb, "VAL_AdjFactor", f"Val!$D${P['adj']}")
    row("caev", "調整後 Comparable EV＝倍數 × 調整係數 × NTM 營收總額", "$B", s3(lambda s: f"={c('mult')}*{X[s]}{P['adj']}*{X[s]}{P['ntm']}"))
    row("caps", "調整後 Comparable 每股", "$/股", s3(lambda s: f"=MAX(0,({X[s]}{P['caev']}+{c('netcash')})/{c('shares')})"), fmt="#,##0.00", bold=True)
    row("scen", "情境目標價＝DCF × 權重 ＋ 調整後 Comparable ×（1 − 權重）", "$/股", s3(lambda s: f"={I('val_w_dcf')}*{X[s]}{P['dcf']}+(1-{I('val_w_dcf')})*{X[s]}{P['caps']}"), fmt="#,##0.00")
    row("scenup", "情境上檔空間", "比例", s3(lambda s: f"={X[s]}{P['scen']}/{c('px')}-1"), fmt="0.0%")

    # ═══ 7. 敏感度 ═══
    sec("7. 敏感度（龍捲風；一次改一個輸入，整本活頁簿以 LibreOffice 重算後的加權目標價；建置快照，tools/val_sens.py → data/val_sens.yaml）")
    header(ws, r[0], ["案例", "低值情況下加權目標價", "高值情況下加權目標價", "低值設定", "高值設定"])
    r[0] += 1
    s0 = r[0]
    cases = (sens or {}).get("cases", [])
    for cs in cases:
        n = r[0]
        put(ws, f"A{n}", cs["name"], F_CALC)
        put(ws, f"B{n}", cs.get("tp_lo"), F_IN, fmt="#,##0.00")
        put(ws, f"C{n}", cs.get("tp_hi"), F_IN, fmt="#,##0.00")
        put(ws, f"D{n}", cs.get("lo_label", ""), F_NOTE)
        put(ws, f"E{n}", cs.get("hi_label", ""), F_NOTE)
        r[0] += 1
    if not cases:
        put(ws, f"A{r[0]}", "（尚未產生：執行 tools/val_sens.py）", F_NOTE)
        r[0] += 1
    nm(wb, "VAL_SensTbl", f"Val!$A${s0}:$C${max(s0, r[0] - 1)}")
    row("sbase", "Sens 快照的基準加權目標價（建置時）", "$/股", [None], fmt="#,##0.00")
    put(ws, f"B{P['sbase']}", (sens or {}).get("base_tp"), F_IN, fmt="#,##0.00")
    put(ws, f"C{P['sbase']}", f"=VAL_TP_W", fmt="#,##0.00")
    put(ws, f"K{P['sbase']}", f"B＝快照（{(sens or {}).get('model', '未產生')}）；C＝活公式；不同即過期（Checks WARN）", F_NOTE)
    nm(wb, "VAL_SensBase", f"Val!$C${P['sbase']}")
    P["sens_rows"] = (s0, r[0] - 1)

    # ═══ 0. 結論回填 ═══
    end = r[0]
    r[0] = concl0
    header(ws, r[0], ["項目", "單位", "數值", "", "", "", "", "", "", "", "說明"])
    r[0] += 1
    W3 = lambda key: f"{I('val_w_a')}*$C${P[key]}+{I('val_w_b')}*$D${P[key]}+{I('val_w_c')}*$E${P[key]}"  # noqa: E731
    row("tpdcf", "DCF 目標價（三情境加權 25／50／25）", "$/股", [f"={W3('dcf')}"], fmt="#,##0.00", name="VAL_TP_DCF")
    row("tpcm", "Comparable 目標價：市場主流（未調整；三情境加權）", "$/股", [f"={W3('cmps')}"], fmt="#,##0.00", name="VAL_TP_CompMain")
    row("tpca", "Comparable 目標價：調整後（FCF 利益率比例；三情境加權；進加權）", "$/股", [f"={W3('caps')}"], fmt="#,##0.00", name="VAL_TP_CompAdj")
    row("tpw", "加權目標價＝DCF × 50% ＋ 調整後 Comparable × 50%", "$/股", [f"={I('val_w_dcf')}*{c('tpdcf')}+(1-{I('val_w_dcf')})*{c('tpca')}"], fmt="#,##0.00",
        name="VAL_TP_W", bold=True)
    row("pxc", "現價（IPO 區間中點）", "$/股", [f"={c('px')}"], fmt="#,##0.00")
    row("upw", "上檔空間：加權目標價 ÷ 現價 − 1", "比例", [f"={c('tpw')}/{c('px')}-1"], fmt="0.0%", name="VAL_Up_W", bold=True)
    row("upcm", "上檔空間：市場主流 Comparable", "比例", [f"={c('tpcm')}/{c('px')}-1"], fmt="0.0%", name="VAL_Up_CompMain")
    row("updcf", "上檔空間：DCF", "比例", [f"={c('tpdcf')}/{c('px')}-1"], fmt="0.0%", name="VAL_Up_DCF")
    u = c("upw")
    rate = (f'=IF({u}>{I("val_r_sb")},"強力買進",IF({u}>{I("val_r_b")},"買進",IF({u}>={I("val_r_n")},"中立",IF({u}>={I("val_r_s")},"放空","強力放空"))))')
    row("rating", "評等", "—", [rate], fmt="@", name="VAL_Rating", bold=True, note="邊界：>+50% 強力買進、>+25% 買進、≥−25% 中立、≥−50% 放空、其餘強力放空")
    row("dist", "距最近評等邊界（百分點）", "pt", [f"=MIN(ABS({u}-{I('val_r_sb')}),ABS({u}-{I('val_r_b')}),ABS({u}-{I('val_r_n')}),ABS({u}-{I('val_r_s')}))*100"], fmt="0.0")
    row("near", "距邊界 5pt 內（1＝是；原則 36 揭露）", "旗標", [f"=IF({c('dist')}<={I('val_r_near')}*100,1,0)"], fmt="0")
    note_f = ('="評等規則：加權目標價上檔空間 >+50% 強力買進、>+25% 買進、≥−25% 中立、≥−50% 放空、其餘強力放空。本次上檔空間 "&TEXT(' + u + '*100,"0.0")&"%（加權目標價 "&TEXT('
              + c("tpw") + ',"0.00")&"／現價 "&TEXT(' + c("px") + ',"0.00")&"），評等"&' + c("rating") + '&"；距最近邊界 "&TEXT(' + c("dist") + ',"0.0")&"pt"&IF('
              + c("near") + '=1,"（5pt 內：評等對假設敏感，揭露）","（不在 5pt 內）")')
    row("rnote", "評等說明（含距邊界揭露）", "—", [note_f], fmt="@", name="VAL_RatingNote")
    row("pricez", "IPO 價格區間 70／78 下的上檔空間（加權目標價不變的近似；精確值見 Sens）", "比例", [f"={c('tpw')}/{c('pxlo')}-1", f"={c('tpw')}/{c('pxhi')}-1"], fmt="0.0%")
    row("meth", "方法", "—", ["DCF（FCFF；EBIT＝Funding F06 現金口徑、股權報酬不加回；稅 21% 扣自身 NOL；WACC≈ke；終值 RONIC＝WACC、g 3%）50%＋調整後 Comparable"
                              "（同業 EV／NTM 營收中位數 × MIN(1, FY2030 自由現金流率 ÷ 同業 FCF 利益率中位數)）50%；情境 A／B／C 25／50／25；現價＝IPO 區間中點 74（J8）；"
                              "判斷題 J5 DCF＋Comparable、J6 FCF 利益率調整、J7 API 任務數高／低端＋Q4 讀法、J8 區間中點"], fmt="@", name="VAL_Method")
    row("wsum", "檢查：情境權重合計 − 1", "—", [f"={I('val_w_a')}+{I('val_w_b')}+{I('val_w_c')}-1"], fmt="0.000000")
    row("tvmax", "檢查：三情境終值占 EV 最大值", "比例", [f"=MAX($C${P['tvsh']}:$E${P['tvsh']})"], fmt="0.0%")
    assert r[0] <= concl0 + 22, r[0]
    r[0] = end
    ws.column_dimensions["A"].width = 70
    ws.column_dimensions["B"].width = 10
    for col in DC:
        ws.column_dimensions[col].width = 12
    ws.column_dimensions["K"].width = 80
    ws.freeze_panes = "C5"
    ctx.P["a11"] = dict(P=P, snap=snap, sens=sens)


def checks(ctx):
    P = ctx.P["a11"]["P"]
    snap, sens = ctx.P["a11"]["snap"], ctx.P["a11"]["sens"]
    rows = [
        ("v_wsum", "v0.2.0 Val：情境權重合計 − 1", f"=Val!$C${P['wsum']}", 0, "tol", ""),
        ("v_tp", "v0.2.0 Val：加權目標價（$／股）", "=VAL_TP_W", None, "info", "DCF 50%＋調整後 Comparable 50%"),
        ("v_rating", "v0.2.0 Val：評等", "=VAL_Rating", None, "info", ""),
        ("v_tv", "v0.2.0 Val：三情境終值占 EV 最大值 > 80% 的超出量", f"=MAX(Val!$C${P['tvmax']}-0.8,0)", "0", "warn", "個股DCF標準規格 第 4 節：>80% 在報告標示"),
    ]
    # 快照過期：B 快照 vs 活公式 B（營收淨額、EBIT 合計差）
    b = (snap or {}).get("scenarios", {}).get("B", {}).get("values")
    if b:
        ref = sum(b["rev"]) + sum(b["ebit"])
        rows.append(("v_snap", "v0.2.0 Val：情境快照過期檢查＝|活公式 B（營收淨額＋EBIT，FY2026–2030 合計）− 快照 B|", f"=ABS(SUM(Val!$C${P['B_rev']}:$G${P['B_rev']})+SUM(Val!$C${P['B_ebit']}:$G${P['B_ebit']})-{ref!r})",
                     "0.000001", "warn", "WARN＝A／C 快照需以 tools/val_snapshots.py 重產"))
    else:
        rows.append(("v_snap", "v0.2.0 Val：情境快照過期檢查（快照未產生）", "=1", "0.000001", "warn", "WARN＝A／C 快照需以 tools/val_snapshots.py 重產"))
    if sens and sens.get("base_tp") is not None:
        rows.append(("v_sens", "v0.2.0 Val：Sens 快照過期檢查＝|活公式加權目標價 − Sens 快照基準|", f"=ABS(VAL_TP_W-Val!$B${P['sbase']})", "0.000001", "warn",
                     "WARN＝以 tools/val_sens.py 重產"))
    else:
        rows.append(("v_sens", "v0.2.0 Val：Sens 快照過期檢查（未產生）", "=1", "0.000001", "warn", "WARN＝以 tools/val_sens.py 重產"))
    return rows
