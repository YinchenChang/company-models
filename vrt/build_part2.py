# build_vrt.py 的後半（由 build_vrt.py 以 exec 載入，共用同一命名空間）：DCF、Val、Sens、Coord、Recon、RevDCF、README、Sources、存檔
FY_X = list(range(2027, 2033))
DCFW = wb.create_sheet("DCF")
DCFW["A1"] = ("DCF — 標準版：無槓桿自由現金流 FCFF＝NOPAT＋折舊攤銷 − 資本支出 − 營運資金增加；WACC 折現到評價日 2026-12-31（期中慣例；FY2027 起整年）；"
              "終值＝NOPAT ×（1＋g）×（1 − g ÷ RONIC）÷（WACC − g）。共用模組 tools/fcff_dcf.py；規格 docs/plan/個股DCF標準規格.md v1.0")
DCFW["A2"] = ("EBIT 取 PL 頁各情境「營業利益」（GAAP；FY2030–FY2032 為收斂法）；NOPAT＝EBIT ×（1 − tax_f）。Non-GAAP 橋接（原則 40）：公司調整後營業利益加回攤銷與併購項目，"
              "本模型 DCF 用 GAAP 營業利益，攤銷已在折舊攤銷（fc_da）加回，股權報酬視為費用、不加回。折舊攤銷、資本支出、營運資金比例見 In 頁；FY2029 起資本支出比例線性收斂到穩態。"
              "淨現金由資產負債表日 2026-09-30 滾到評價日（stub：4Q26 FCFF 用 Q 頁模型季度數，扣已宣告股利）。UtilityInnovation 收購（2026-10-26 交割，現金 14.5 億美元）基準視為價值中性。")
STUB_Q4 = "(Q!$F$8*(1-tax_f)+Q!$F$4*(fc_da-fc_cx-fc_cx_d)-fc_wc*(Q!$G$4-Q!$B$4)*stub_frac)"
DP = [dict(name=s, label={"A": "情境 A（樂觀）", "B": "情境 B（基準）", "C": "情境 C（保守）"}[s] + "：" + SCN_T[s].split("：")[1],
           rev={y: f"PL!{FC[y]}{PB[s]['rev']}" for y in FY_X}, rev0=f"PL!{FC[2026]}{PB[s]['rev']}",
           ebit={y: f"PL!{FC[y]}{PB[s]['oi']}" for y in FY_X}, tax={y: "tax_f" for y in FY_X}, ebit_note="PL 頁營業利益（GAAP）", stub_fcff=STUB_Q4) for s in SC]
RD_ = FD.write_detail(DCFW, 4, DP, FY_X, 1.0, conv_from=2029, title="三情境 FCFF 明細（US$B；會計年度＝日曆年，FY2027 起整年）", fmt="#,##0.000", stub=True)
RD_["end"] = FD.write_checks(DCFW, RD_["end"], nc_bs="3.5129+0.1988-2.9285-0.3397",
                             nc_note="2026-09-30 資產負債表（US$B）：現金 3.5129＋短期投資 0.1988 − 債務 2.9285 − 長期租賃 0.3397 − 少數股權 0（Q3 新聞稿與 10-Q）")
DCFW.column_dimensions["A"].width = 60; DCFW.column_dimensions["B"].width = 40
for _c in "CDEFGH": DCFW.column_dimensions[_c].width = 13

# ---------------- Val ----------------
V = wb.create_sheet("Val")
T = {}
V["A1"] = "Val — 估值（美元）：目標價＝DCF 與 Comparable（相對倍數調整後）加權；市場隱含倍數、反向檢驗見 RevDCF 頁"
V["A2"] = ("Comparable 同業：Eaton、Schneider Electric、台達電（機房電力與電氣同業；Vertiv 自身列為本公司不入同業中位數）。前瞻 P/E＝2026-11-13 收盤 ÷ 資料商前瞻 EPS（同 delta 模型 Val 頁同業表，stockanalysis，2026-10）。"
           "目標價看本模型 GAAP EPS，同業為資料商（混合口徑）前瞻 EPS，口徑差異見 Recon。")
V.cell(4, 1, "1. 現價與市場隱含倍數").font = BOLD
hdr(V, 5, ["項目", "A", "B", "C", "說明"])
V.cell(6, 1, "現價（2026-11-13 收盤）"); V["B6"] = "=px"; V["B6"].number_format = "#,##0.00"
for rr, lab, y in ((7, "現價 ÷ FY2026 EPS（錨點，GAAP）", 2026), (8, "現價 ÷ FY2027 EPS（模型）", 2027), (9, "現價 ÷ FY2028 EPS（模型）", 2028), (10, "現價 ÷ FY2029 EPS（模型）", 2029)):
    V.cell(rr, 1, lab)
    for j, s in enumerate(SC):
        c = "BCD"[j]; V[f"{c}{rr}"] = f"=px/PL!{FC[y]}{SUMR[(s, 'eps')]}"; V[f"{c}{rr}"].number_format = "0.0"
V["E8"] = "FY2027 三情境毛利率與市場倍數相同（情境倍數從 FY2028 開始），EPS 因毛利率加減而略有差異"
t0 = 13; T["t0"] = t0
V.cell(t0, 1, "0. 目標價（DCF 與 Comparable 加權；三情境欄 B／C／D＝A 樂觀／B 基準／C 保守）").font = BOLD
SDCF = FD.write_summary(wb, V, t0 + 1, RD_, list(SC), ["tp_wA", "tp_wB", "tp_wC"], fmt="#,##0.00",
                        title="0.1 DCF（標準版：無槓桿自由現金流 FCFF＋WACC；明細見 DCF 頁；B／C／D＝A 樂觀／B 基準／C 保守）", beta_note="Yahoo 原始 β 經 Blume 調整")
for _k in ("pvx", "pvtv", "ev", "nc", "stub", "ncr"):
    for _c in "BCD": V[f"{_c}{SDCF[_k]}"].number_format = "#,##0.000"
T.update({k: v for k, v in SDCF.items()})
r_ = t0 + 23
def tset(r, key, label, fs, fmt, note=""):
    V.cell(r, 1, label); T[key] = r
    for c, f in zip("BCD", fs):
        if f is None: continue
        V[f"{c}{r}"] = f; V[f"{c}{r}"].number_format = fmt
    if note: V[f"E{r}"] = note
V.cell(r_, 1, "0.2 Comparable（機房電力與電氣同業前瞻 P/E）").font = BOLD; r_ += 1
hdr(V, r_, ["同業與前瞻 P/E", "代號", "前瞻 12 月 P/E", "日期", "來源", "標記"]); T["ph"] = r_; r_ += 1
P1 = r_
for nm, tk, pe, src in (("Eaton", "ETN", "=352.97/(376.52/24.81)", "stockanalysis 統計頁 2026-10-27：價格 376.52 ÷ Forward PE 24.81（同 delta 模型）；2026-11-13 收盤 352.97"),
                        ("Schneider Electric", "SU.PA", "=243.45/(241/19.46)", "stockanalysis 統計頁 2026-10-27：價格 241.00 歐元 ÷ Forward PE 19.46（同 delta 模型）；2026-11-13 收盤 243.45 歐元"),
                        ("台達電（Delta Electronics）", "2308.TW", "=1240/(1560/30.51)", "stockanalysis 個股頁 2026-10-22：價格 1,560 ÷ Forward PE 30.51（同 delta 模型）；2026-11-13 收盤 1,240 新台幣")):
    V.cell(r_, 1, nm); V.cell(r_, 2, tk); cc = V.cell(r_, 3, pe); cc.font = BLUE; cc.number_format = "0.0"
    V.cell(r_, 4, "2026-11-13"); V.cell(r_, 5, "P/E＝2026-11-13 收盤（價格檔）÷ 資料商前瞻 EPS；" + src); V.cell(r_, 6, "Interested-party（賣方共識）"); r_ += 1
P2 = r_ - 1; T["p1"], T["p2"] = P1, P2
tset(r_, "pmed", "同業前瞻 P/E 中位數", [f"=MEDIAN(C{P1}:C{P2})", None, None], "0.0", "Eaton、Schneider Electric、台達電；中國 A 股不列（結構性溢價，Andy 2026-10-25）；Vertiv 自身不入中位數"); r_ += 1
tset(r_, "feps", "12 個月後的前瞻 EPS（情境加權）", ["=" + "+".join(f"tp_w{s}*(0.16666666666666666*PL!{FC[2027]}{PB[s]['eps']}+0.8333333333333334*PL!{FC[2028]}{PB[s]['eps']})" for s in SC), None, None], "0.00",
     "1/6 × FY2027 EPS＋5/6 × FY2028 EPS（與 delta 同權重；評價日 2026-12-31 後 12 個月）"); r_ += 1
tset(r_, "comptp", "Comparable 目標價", [f"=B{T['pmed']}*B{T['feps']}", None, None], "#,##0.00", "同業中位數 × 上列 EPS"); r_ += 2
V.cell(r_, 1, "0.3 結論").font = BOLD; r_ += 1
hdr(V, r_, ["項目", "數值", "上檔空間", "", "說明"]); T["h3"] = r_; r_ += 1
T["c0"] = r_
tset(r_, "dcftp", "DCF 目標價", [f"=B{SDCF['dcf']}", f"=B{r_}/px-1", None], "#,##0.00", "引用 0.1：三情境價值推到 12 個月後的加權"); V[f"C{r_}"].number_format = "+0.0%;-0.0%;0.0%"; r_ += 1
tset(r_, "compadj", "Comparable（調整後，進加權）", ["=tp_comp_adj", f"=B{r_}/px-1", None], "#,##0.00", "同業中位數 × EPS × 相對倍數調整係數（0.5 節）"); V[f"C{r_}"].number_format = "+0.0%;-0.0%;0.0%"; r_ += 1
tset(r_, "comp2", "Comparable 目標價（市場主流同業法）", [f"=B{T['comptp']}", f"=B{r_}/px-1", None], "#,##0.00", "市場主流同業法"); V[f"C{r_}"].number_format = "+0.0%;-0.0%;0.0%"; r_ += 1
tset(r_, "tpw", "目標價（DCF 與 Comparable 加權）", [f"=tp_w_dcf*B{T['dcftp']}+(1-tp_w_dcf)*tp_comp_adj", f"=B{r_}/px-1", None], "#,##0.00", "DCF 權重 × DCF 目標價＋（1 − DCF 權重）× Comparable（調整後）；權重見 In 頁")
V[f"C{r_}"].number_format = "+0.0%;-0.0%;0.0%"; r_ += 1
tset(r_, "cons", "市場共識目標價", ["=cons_tp", f"=B{r_}/px-1", None], "#,##0.00", "Yahoo 2026-11-17（30 位分析師）；只作對帳，不作校準"); V[f"C{r_}"].number_format = "+0.0%;-0.0%;0.0%"; r_ += 1
tset(r_, "pxr", "現價（2026-11-13 收盤）", ["=px", None, None], "#,##0.00", "Yahoo Finance 日線收盤（價格檔）"); r_ += 1
T["c1"] = r_ - 1
tset(r_, "upw", "上檔空間：加權目標價", [f"=B{T['tpw']}/px-1", None, None], "0.0%", "相對現價"); r_ += 1
tset(r_, "upc", "上檔空間：Comparable 目標價（未調整）", [f"=B{T['comp2']}/px-1", None, None], "0.0%"); r_ += 1
tset(r_, "geps", "與市場共識的差距：EPS 因素（本模型 FY2027 GAAP EPS ÷ 共識調整後 EPS − 1）", [f"=PL!{FC[2027]}{PB['B']['eps']}/cons_eps27-1", None, None], "0.0%", "口徑不同（共識為 Non-GAAP，約加回攤銷與調整項目），差距含口徑差"); r_ += 1
tset(r_, "band", "價格帶（DCF 保守～樂觀；推到 12 個月後）", [f'=TEXT(D{SDCF["v12"]},"#,##0")&"～"&TEXT(B{SDCF["v12"]},"#,##0")', None, None], "@", "C 情境價值～A 情境價值"); r_ += 1
V.cell(r_, 1, "目標價方法"); V[f"B{r_}"] = "加權目標價＝DCF（標準版 FCFF＋WACC，三情境加權）與 Comparable（機房電力同業前瞻 P/E 中位數 × 本公司 FY2027–FY2028 EPS × 10 年相對倍數調整係數）各半；投資建議依上檔空間固定規則。"; T["m"] = r_; r_ += 1
for k_ in ("dcftp", "compadj", "tpw"): V[f"B{T[k_]}"].fill = KEY
V.cell(T["tpw"], 1).font = BOLD
defname("tp_final", f"Val!$B${T['tpw']}"); defname("tp_dcf", f"Val!$B${T['dcftp']}"); defname("tp_comp", f"Val!$B${T['comptp']}")
V["A3"] = f"目標價：見第 0 節（本頁第 {t0} 列起；DCF 與 Comparable 加權）"
V.column_dimensions["A"].width = 62
for c in "BCD": V.column_dimensions[c].width = 16
V.column_dimensions["E"].width = 80
# 相對倍數調整（0.5 節）
RMR = RMX.add_block(wb, V, own="VRT", peers=[("Eaton", "ETN"), ("Schneider Electric", "SU.PA"), ("台達電（Delta Electronics）", "2308.TW")], metric="PE",
                    comp_ref=f"Val!$B${T['comptp']}", pxfmt="#,##0.00", own_name="Vertiv")
T["rm"] = RMR
# 上檔空間級距與投資建議
GR = V.max_row + 3
_U = f"B{T['upw']}"
V.cell(GR - 1, 1, "0.4 上檔空間級距與投資建議（依「上檔空間：加權目標價」；Andy 2026-10-25 固定規則）").font = BOLD
tset(GR, "grade", "上檔空間級距", [f'=IF({_U}>0.5,"大於 50%",IF({_U}>0.25,"25%–50%",IF({_U}>=0,"0%–25%",IF({_U}>=-0.25,"0%～−25%",IF({_U}>=-0.5,"−25%～−50%","小於 −50%")))))', None, None], "General", "分六級；負值區間同樣重要")
tset(GR + 1, "rec", "投資建議", [f'=IF({_U}>0.5,"強力買進",IF({_U}>0.25,"買進",IF({_U}>=-0.25,"中立",IF({_U}>=-0.5,"放空","強力放空"))))', None, None], "General",
     "大於 +50% 強力買進、+25%～+50% 買進、−25%～+25% 中立、−50%～−25% 放空、小於 −50% 強力放空（剛好 +25% 為中立）")
V.cell(GR + 2, 1, "估值評等（與投資建議相同；自由現金流為正，收支關卡不適用）"); V[f"B{GR + 2}"] = f"=B{GR + 1}"; T["rate"] = GR + 2
REC = f"Val!B{T['rec']}"

# ---------------- Sens ----------------
SE = wb.create_sheet("Sens")
SE["A1"] = ("Sens — 敏感度。第一部分：目標價參數（標準版 DCF 的折現與資本參數，閉式公式）；第二部分：業務主因（每個案例改一個輸入後整本重建、LibreOffice 重算，快照 vrt/sens.json；"
            "In 頁改動後須重跑 vrt/sens_vrt.py，否則模型狀態顯示 STALE）。加權目標價一次動一個輸入，Comparable 不變（DCF 權重案例除外）。")
SE["A2"] = "讀法：相對基準＝案例加權目標價 ÷ 基準加權目標價 − 1。"
TPS0 = 3
_tvs = FD.write_sens(SE, TPS0, RD_, list(SC), ["tp_wA", "tp_wB", "tp_wC"], comp_ref="tp_comp_adj", base_ref=f"Val!$B${T['tpw']}", fmt="#,##0.00",
                     title="目標價參數敏感度（標準版 DCF）：加權目標價一次動一個參數；Comparable 不變（DCF 權重案例除外）", beta_u=None, w_dcf_lohi=True)
r_ = _tvs[1] + 1
# 自訂案例：UtilityInnovation 收購現金
SE.cell(r_ + 1, 1, "UtilityInnovation 收購現金（高）"); SE.cell(r_ + 1, 2, "收購現金 14.5 億美元計為純流出、不計被收購業務獲利（下緣）；基準＝價值中性"); SE.cell(r_ + 1, 8, "uig_adj")
SE.cell(r_ + 1, 3, "—"); SE.cell(r_ + 1, 4, "—")
_dcfx = "+".join(f"{w}*{FD.value_expr(RD_, n, nc='(dcf_nc-uig_adj_hi)')}" for w, n in zip(["tp_wA", "tp_wB", "tp_wC"], SC))
SE.cell(r_ + 1, 5, f"=tp_w_dcf*({_dcfx})+(1-tp_w_dcf)*tp_comp_adj").number_format = "#,##0.00"
SE.cell(r_ + 1, 6, f"=E{r_ + 1}/G{r_ + 1}-1").number_format = "+0.0%;-0.0%;0.0%"
SE.cell(r_ + 1, 7, f"=Val!$B${T['tpw']}").number_format = "#,##0.00"
SE_FD_LAST = r_ + 1
# 業務主因表
BIZ = [("dcp_26", "資料中心設施產品占 FY2026 營收（份額基準）", "%"), ("rel_27", "Vertiv 份額 FY2027 相對值", "x"), ("rel_28", "Vertiv 份額 FY2028 相對值", "x"), ("rel_29", "Vertiv 份額 FY2029 相對值", "x"),
       ("mk_adj", "設施市場 2028–2029 倍數（Wattonomics 風險）", "x"), ("g_oth", "其餘營收年增（服務、備品）", "%"), ("gm_x", "毛利率相對 FY2026 變動", "pt"), ("phi", "營業費用彈性", "x"),
       ("opx_adj", "營業費用額外占營收比", "pt"), ("nop_f", "業外淨額", "US$B"), ("tax_f", "有效稅率", "%"), ("sh_g", "稀釋股數年變動", "%"), ("q4_rev", "4Q26 營收（指引區間）", "US$B"),
       ("q_gm", "4Q26 毛利率", "%"), ("q_opx", "4Q26 營業費用率", "%"), ("q_nop", "4Q26 業外淨額", "US$B"), ("q_tax", "4Q26 稅率", "%"), ("q_adj", "4Q26 調整項目（Non-GAAP 橋接，不影響 EPS 與 DCF）", "US$B"),
       ("opm_f", "營業利益率收斂係數", "x"), ("tp_g", "收斂基準成長率", "%")]
BR0 = r_ + 4
SE.cell(BR0, 1, "業務主因（整本重建；一次動一個輸入至 In 頁低／高值；快照 vrt/sens.json）").font = BOLD
for j, h in enumerate(["參數", "設定（低／高）", "低：加權目標價", "低：相對基準", "高：加權目標價", "高：相對基準", "基準加權目標價", "鍵", "低：DCF 目標價", "高：DCF 目標價"], 1):
    c = SE.cell(BR0 + 1, j, h); c.font = BOLD; c.fill = HDR
SNAP = None
if os.path.exists(os.path.join(HERE, "sens.json")) and not os.environ.get("VRT_SENS_RUN"):
    SNAP = json.load(open(os.path.join(HERE, "sens.json"), encoding="utf-8"))
rr = BR0 + 2; BIZR = {}
for key, lab, unit in BIZ:
    SE.cell(rr, 1, lab); SE.cell(rr, 2, f'="{key}：低 "&TEXT({key}_lo,"0.000")&"／高 "&TEXT({key}_hi,"0.000")&"（基準 "&TEXT({key},"0.000")&"）"'); SE.cell(rr, 8, key)
    for col, side in ((3, "低"), (5, "高")):
        v = SNAP["biz"].get(key, {}).get(side) if SNAP else None
        if v:
            SE.cell(rr, col, v["tp"]).number_format = "#,##0.00"
            SE.cell(rr, col + 1, f"={'CE'[(col - 3) // 2]}{rr}/G{rr}-1").number_format = "+0.0%;-0.0%;0.0%"
            SE.cell(rr, 9 if side == "低" else 10, v["dcf"]).number_format = "#,##0.00"
        else:
            SE.cell(rr, col, "—"); SE.cell(rr, col + 1, "—")
    SE.cell(rr, 7, f"=Val!$B${T['tpw']}").number_format = "#,##0.00"
    BIZR[key] = rr; rr += 1
BIZ_LAST = rr - 1
SE.cell(rr, 1, "註：業務主因案例為快照值（整本重算），非即時公式；基準欄為即時值，快照基準見下方；兩者不同＝模型狀態 STALE")
rr += 2
SE.cell(rr, 1, "豁免（不列案例的鍵）").font = BOLD; rr += 1
EXEMPT = ["dcp_25：不列：只影響 FY2025 歷史份額的顯示（Units 頁），不進任何預測與目標價",
          "mk_dist_26：不列：Wattonomics 市場輸入為單點 Derived 值（無低／高區間），市場風險由 mk_adj 與三情境倍數涵蓋"]
EXEMPT_ROWS = []
for t in EXEMPT:
    SE.cell(rr, 1, t); EXEMPT_ROWS.append(rr); rr += 1
rr += 1
import hashlib as _h
BUILD_ID = "vrt-v" + VER + "-" + _h.sha1((open(os.path.join(HERE, "build_vrt.py"), "rb").read() + open(os.path.join(HERE, "build_part2.py"), "rb").read())).hexdigest()[:10]
SE.cell(rr, 1, "build_id"); SE.cell(rr, 4, BUILD_ID); BID_ROW = rr; rr += 1
SE.cell(rr, 1, "輸入指紋（本檔即時重算）"); SE.cell(rr, 4, f"=SUM(In!$D$4:$F${INLAST})"); SE.cell(rr, 4).number_format = "0.000000"; FP_ROW = rr; rr += 1
SE.cell(rr, 1, "輸入指紋（快照；整本重算時）"); SE.cell(rr, 4, SNAP["_base"]["fp"] if SNAP else None); SE.cell(rr, 4).number_format = "0.000000"; FPS_ROW = rr; rr += 1
SE.cell(rr, 1, "模型狀態"); SE.cell(rr, 4, f'=IF(ISNUMBER(D{FPS_ROW}),IF(ABS(D{FP_ROW}-D{FPS_ROW})<0.000001,"CURRENT","STALE"),"STALE")'); STATE_ROW = rr; rr += 1
SE.cell(rr, 1, "快照基準加權目標價（整本重算）"); SE.cell(rr, 4, SNAP["_base"]["tp"] if SNAP else None); SE.cell(rr, 4).number_format = "#,##0.00"; SNAPTP_ROW = rr; rr += 1
SE.cell(rr, 1, "檢查：快照基準 − 即時加權目標價（應為 0）"); SE.cell(rr, 4, f"=IF(ISNUMBER(D{SNAPTP_ROW}),D{SNAPTP_ROW}-Val!$B${T['tpw']},0)"); SE.cell(rr, 4).number_format = "0.000000"; rr += 1
SE.column_dimensions["A"].width = 50; SE.column_dimensions["B"].width = 62
for c in "CDEFGIJ": SE.column_dimensions[c].width = 16

# ---------------- Coord：三層協調（Vertiv 讀 Wattonomics 層） ----------------
CO = wb.create_sheet("Coord")
CO["A1"] = "Coord — 三層協調：Vertiv 的機房設施市場讀 Wattonomics v0.3.6（不讀 Siliconomics 的電源與散熱）；公司實績對層級隱含值並列。"
CO["A2"] = ("對帳狀態：本模型讀 Wattonomics 層，未發現 Wattonomics 快照有「對帳中」旗標（names 內只有 NV 對帳項 WIF_Recon_NV_*，不含設施層），故不建 co_flag；"
            "另一側讀數：不適用（電力與機房範圍定案 2026-11-17）：機房設施供應商只讀 Wattonomics 廠房拆分，該數字只有 Wattonomics 一個權威層、沒有 Siliconomics 側的另一讀數，三層協調 C4 兩讀數不適用；與 Wattonomics 的差異走對帳單。")
hdr(CO, 4, ["讀取的層級與用途", "層級", "版本", "用在哪裡"])
CO["A5"] = "機房設施支出（配電、UPS、冷卻）"; CO["B5"] = "Wattonomics"; CO["C5"] = "v0.3.6"; CO["D5"] = "Units 頁可及市場 → PL 頁資料中心設施產品營收（FY2027–FY2029）"
CO["A6"] = "快照檔"; CO["B6"] = "wattonomics/data/wattonomics_snapshot_v0.3.6.json"; CO["C6"] = WSHA; CO["D6"] = "names.WIF_Spend_fac_<dist|ups|cool|gen|sum>_<年>"
CO["A7"] = "Siliconomics"; CO["B7"] = "不讀"; CO["C7"] = "v0.20（RECO_FLAG 0/1/1/1，2026–2029）"; CO["D7"] = "Siliconomics 的電源與散熱變數為機櫃內（台達電）口徑，不含 Vertiv 的機櫃外設施內容"
hdr(CO, 10, ["公司實績對層級隱含值", "FY2025", "FY2026", "說明", "", "單位／口徑", "來源與等級"])
CO["A11"] = "Vertiv 資料中心設施產品營收（推導）"; CO["B11"] = "=Units!C11"; CO["C11"] = "=Units!D11"; CO["F11"] = "US$B；營收 × 資料中心設施產品占比（Assumed）"; CO["G11"] = "Q 頁營收（Verified）× In 頁 dcp（Assumed）"
CO["A12"] = "Wattonomics 可及市場（配電＋UPS＋冷卻）"; CO["B12"] = "=Units!C8"; CO["C12"] = "=Units!D8"; CO["F12"] = "US$B；機櫃外設施支出"; CO["G12"] = "Wattonomics v0.3.6（Derived）"
CO["A13"] = "Vertiv 份額（÷ 可及市場）"; CO["B13"] = "=B11/B12"; CO["C13"] = "=C11/C12"; CO["F13"] = "%"; CO["G13"] = "Derived"
CO["A14"] = "份額變動（FY2026 ÷ FY2025 − 1）"; CO["C14"] = "=C13/B13-1"; CO["F14"] = "%"; CO["G14"] = "Derived"
CO["A15"] = "Vertiv 營收年增"; CO["C15"] = "=a26_rev/a25_rev-1"; CO["F15"] = "%"; CO["G15"] = "Verified／Interested-party（4Q26 指引）"
CO["A16"] = "Wattonomics 可及市場年增"; CO["C16"] = "=C12/B12-1"; CO["F16"] = "%"; CO["G16"] = "Derived"
for c in ("B11", "C11", "B12", "C12"): CO[c].number_format = "#,##0.00"
for c in ("B13", "C13", "C14", "C15", "C16"): CO[c].number_format = "0.0%"
CO["A18"] = ('="解讀：Wattonomics 可及市場 FY2026 成長 "&TEXT(C16,"0%")&"，Vertiv 營收成長 "&TEXT(C15,"0%")&"，Vertiv 份額由 "&TEXT(B13,"0.0%")&" 降到 "&TEXT(C13,"0.0%")&"；'
             '市場以裝機年度計支出，Vertiv 於出貨認列，backlog 使營收落後市場，差距不判定為層級錯誤，列入待判斷。"')
CO["A19"] = "Wattonomics 2030–2032 的新增容量下降（可及市場轉降），本模型 FY2030 起不讀市場，改用收斂法；建議 Wattonomics 補充設施端的更換與服務需求，避免市場層在 2030 年後系統性偏低。"
CO.column_dimensions["A"].width = 52; CO.column_dimensions["B"].width = 30; CO.column_dimensions["C"].width = 26; CO.column_dimensions["D"].width = 60
CO.column_dimensions["F"].width = 40; CO.column_dimensions["G"].width = 50

# ---------------- RevDCF ----------------
RV = wb.create_sheet("RevDCF")
RV["A1"] = "RevDCF — 反向檢驗：現價隱含的營收規模（情境 B 路徑；FY2027 起各年營收與現金流同乘一個倍數，其餘假設不變；閉式解，stub 期與淨現金視為不隨倍數縮放）"
RV["A2"] = "現價隱含企業價值＝現價 × 稀釋股數 − 淨現金（滾到評價日）；倍數＝現價隱含企業價值 ÷ 模型情境 B 企業價值。"
RV["A4"] = "現金流折現法（情境 B 路徑）：現價隱含值與模型基準"
hdr(RV, 5, ["項目", "現價隱含", "模型基準"])
RV["A6"] = "企業價值（US$B）"; RV["B6"] = f"=px*sh_v-Val!C{SDCF['ncr']}"; RV["C6"] = f"=Val!C{SDCF['ev']}"
RV["A7"] = "營收規模倍數（FY2027 起各年營收同乘；模型基準＝1）"; RV["B7"] = "=B6/C6"; RV["C7"] = 1
RV["A8"] = "FY2027 營收（US$B）"; RV["B8"] = f"=B7*PL!{FC[2027]}{PB['B']['rev']}"; RV["C8"] = f"=PL!{FC[2027]}{PB['B']['rev']}"
RV["A9"] = "FY2029 營收（US$B）"; RV["B9"] = f"=B7*PL!{FC[2029]}{PB['B']['rev']}"; RV["C9"] = f"=PL!{FC[2029]}{PB['B']['rev']}"
RV["A10"] = "FY2032 營收（US$B）"; RV["B10"] = f"=B7*PL!{FC[2032]}{PB['B']['rev']}"; RV["C10"] = f"=PL!{FC[2032]}{PB['B']['rev']}"
RV["A11"] = "FY2027 營收年增"; RV["B11"] = f"=B8/a26_rev-1"; RV["C11"] = f"=C8/a26_rev-1"
RV["A12"] = "FY2027 營收 ÷ 共識營收"; RV["B12"] = "=B8/cons_rev27"; RV["C12"] = "=C8/cons_rev27"
for rr_ in (6, 8, 9, 10): RV[f"B{rr_}"].number_format = "#,##0.00"; RV[f"C{rr_}"].number_format = "#,##0.00"
RV["B7"].number_format = "0.00"; RV["C7"].number_format = "0.00"
for rr_ in (11, 12): RV[f"B{rr_}"].number_format = "0.0%"; RV[f"C{rr_}"].number_format = "0.0%"
RV["A14"] = ('="現價隱含營收規模約為模型情境 B 的 "&TEXT(B7,"0.00")&" 倍（FY2027 營收 "&TEXT(B8,"0.0")&" 十億美元，年增 "&TEXT(B11,"0%")&"；共識 FY2027 營收 "&TEXT(cons_rev27,"0.0")&" 十億美元）；'
             'Wattonomics 基準在 FY2028 後市場持平，隱含倍數大於 1 表示現價已預期市場或份額高於 Wattonomics 基準。"')
RV.column_dimensions["A"].width = 60; RV.column_dimensions["B"].width = 18; RV.column_dimensions["C"].width = 18

# ---------------- Recon ----------------
RC = wb.create_sheet("Recon")
RC["A1"] = "Recon — 對帳：本模型對公司指引、市場共識與 Yahoo 差異。共識與指引只作對照，不作校準；共識為 Non-GAAP（調整後），本模型為 GAAP，口徑差已標示。"
hdr(RC, 3, ["對照項目", "本模型", "對照值", "差距", "單位", "口徑", "來源", "等級", "", "說明"])
RECS = [("FY2026 營收 vs 公司指引中位", "=a26_rev", 14.65, "US$B", "GAAP 營收", "Vertiv 3Q26 新聞稿：FY2026 營收 14,450–14,850", "Interested-party", "管理層指引；3Q26 後 4Q26 區間為 4,313–4,513，錨點已使用 4Q 指引中位"),
        ("FY2026 營收 vs 共識", "=a26_rev", 14.72, "US$B", "GAAP 營收", "Yahoo 2026-11-17：2026-12 營收 14.72B", "Interested-party", "賣方共識"),
        ("FY2026 GAAP EPS vs 公司 GAAP 指引中位", "=a26_eps", 6.29, "US$", "GAAP 稀釋 EPS", "Vertiv 3Q26 新聞稿：調整後 EPS 7.10–7.20、GAAP 中位 6.29", "Interested-party", "管理層指引；錨點未用 EPS 指引校準"),
        ("FY2026 調整後營業利益 vs 公司指引中位", "=a26_adj", 3.515, "US$B", "Non-GAAP", "Vertiv 3Q26 新聞稿：3,475–3,555", "Interested-party", "管理層指引"),
        ("FY2027 營收 vs 共識", f"=PL!{FC[2027]}{PB['B']['rev']}", "=cons_rev27", "US$B", "GAAP 營收（情境 B）", "Yahoo 2026-11-17：2027-12 營收 19.54B", "Interested-party", "賣方共識；差距主因＝Wattonomics 基準下設施市場成長低於賣方預期"),
        ("FY2027 EPS vs 共識", f"=PL!{FC[2027]}{PB['B']['eps']}", "=cons_eps27", "US$", "本模型 GAAP；共識 Non-GAAP", "Yahoo 2026-11-17：2027-12 EPS 9.874", "Interested-party", "共識加回攤銷與調整項目，約高於 GAAP 0.7–0.9 美元；口徑差未剔除"),
        ("加權目標價 vs 市場共識目標價", f"=Val!B{T['tpw']}", "=cons_tp", "US$", "目標價", "Yahoo 2026-11-17：均值 334.29（30 位）", "Interested-party", "賣方共識；只作對帳"),
        ("Comparable 目標價（未調整）vs 市場共識目標價", f"=Val!B{T['comptp']}", "=cons_tp", "US$", "目標價", "同上", "Interested-party", "—")]
for i, (lab, mdl, ref, unit, kou, src, tag, note) in enumerate(RECS, 4):
    RC.cell(i, 1, lab); RC.cell(i, 2, mdl).number_format = "#,##0.00"; c = RC.cell(i, 3, ref); c.number_format = "#,##0.00"
    if not (isinstance(ref, str)): c.font = BLUE
    RC.cell(i, 4, f"=B{i}/C{i}-1").number_format = "+0.0%;-0.0%;0.0%"; RC.cell(i, 5, unit); RC.cell(i, 6, kou); RC.cell(i, 7, src); RC.cell(i, 8, tag); RC.cell(i, 10, note)
RECN = 3 + len(RECS)
RC.cell(RECN + 2, 1, "Yahoo 與公司原文的差異（公司原文優先；Yahoo 只作交叉核對）").font = BOLD
hdr(RC, RECN + 3, ["項目", "公司原文", "Yahoo", "差距", "單位", "口徑", "來源", "等級", "", "說明"])
for k, (lab, co, ya, note) in enumerate((("FY2023 營業利益", 0.8722, 0.9168, "Yahoo 營業利益口徑與公司 GAAP 營業利益不同"), ("FY2024 營業利益", 1.3674, 1.3820, "同上"), ("FY2025 營業利益", 1.8297, 1.8962, "同上；公司 FY2025 含重組 54.5"),
                                          ("3Q26 營業利益", 0.9838, 0.9911, "公司 983.8；Yahoo 991.1"), ("1Q26 營業利益", 0.4401, 0.4336, "公司為 1H26 − 2Q26 推導；Yahoo 為單季值"), ("FY2025 營收", 10.2299, 10.230, "吻合（四捨五入）"),
                                          ("3Q26 淨利", 0.7645, 0.7645, "吻合"), ("3Q26 總債務（公司：流動 3.4＋長期 2,925.1）", 2.9285, 3.3569, "Yahoo 總債務含租賃負債等；本模型債務與租賃分列")), RECN + 4):
    RC.cell(k, 1, lab); a = RC.cell(k, 2, co); b = RC.cell(k, 3, ya); a.font = BLUE; b.font = BLUE
    a.number_format = "#,##0.000"; b.number_format = "#,##0.000"
    RC.cell(k, 4, f"=C{k}/B{k}-1").number_format = "+0.0%;-0.0%;0.0%"; RC.cell(k, 5, "US$B"); RC.cell(k, 6, "GAAP"); RC.cell(k, 7, "公司新聞稿與 10-Q；Yahoo quoteSummary 2026-11-17（docs/workorders/20261117_新增五家_INTC_GFS_VRT_ABB_TEL_Yahoo基本面快照.md）")
    RC.cell(k, 8, "Verified"); RC.cell(k, 10, note)
RC.column_dimensions["A"].width = 54
for c in "BCDE": RC.column_dimensions[c].width = 14
RC.column_dimensions["G"].width = 70; RC.column_dimensions["J"].width = 70

# ---------------- Sources ----------------
SO = wb.create_sheet("Sources"); SO["A1"] = "來源（公司原文＝Vertiv 新聞稿與 10-Q；Yahoo、賣方與資料商為 Interested-party，只作交叉核對與對帳）"; SO["A1"].font = BOLD
for i, (a, b) in enumerate((
        ("[Verified] Vertiv 3Q26 新聞稿（8-K 附件 99.1，2026-10-28）：營收 4,312.8、營業利益 983.8、調整後營業利益 1,086.9、稀釋 EPS 1.95、4Q26 與 FY2026 指引", "https://www.sec.gov/Archives/edgar/data/0001674101/000162828026067707/q32026exhibit991vrt10282026.htm"),
        ("[Verified] Vertiv 3Q26 10-Q：損益、資產負債表、債務表、現金流量、股利、期後事項（UtilityInnovation 合併協議）", "https://www.sec.gov/Archives/edgar/data/1674101/000162828026067893/vrt-20260930.htm"),
        ("[Verified] Vertiv 2Q26 新聞稿（8-K 附件 99.1，2026-07-29）：2Q26 與 6M26 損益、調整後營業利益、分部營收", "https://www.sec.gov/Archives/edgar/data/0001674101/000162828026050323/q22026exhibit991vrt07292026.htm"),
        ("[Verified] Vertiv 4Q25／FY2025 新聞稿：FY2025 與 FY2024 損益、分部營收、backlog 150 億美元、book-to-bill 約 2.9、現金流量", "https://s205.q4cdn.com/554782763/files/doc_financials/2025/q4/Vertiv_Fourth-Quarter-2025-Earnings-Release.pdf"),
        ("[Verified] Vertiv 4Q23 新聞稿（8-K 附件 99.1，2024-02-21）：FY2023 與 FY2022 損益", "https://www.sec.gov/Archives/edgar/data/1674101/000162828024005765/exhibit991vrt02212024.htm"),
        ("[Derived] Wattonomics v0.3.6 快照：機房設施支出（配電、UPS、冷卻、發電）", "wattonomics/data/wattonomics_snapshot_v0.3.6.json（sha256 " + WSHA + "）"),
        ("[Interested-party] Yahoo quoteSummary 2026-11-17（市值、Beta、共識 EPS 與營收、目標價）；只作交叉核對", "docs/workorders/20261117_新增五家_INTC_GFS_VRT_ABB_TEL_Yahoo基本面快照.md"),
        ("[Verified] 價格：2026-11-13 收盤（Yahoo Finance 日線，主串整理）", "markets/invest_map/data/closes_extra_20261113.csv；closes_20261020_20261113.csv"),
        ("[Interested-party] 同業前瞻 EPS（Eaton、Schneider、台達電）：stockanalysis，同 delta 模型 Val 頁同業表", "delta/dist（Val 頁 0.2 同業表）"),
        ("[Interested-party] UtilityInnovation Group 收購（現金 14.5 億美元＋最多 11.5 億美元 EBITDA 對價；2026-10-26 交割）", "https://www.unite.ai/vertiv-strikes-1-45b-deal-for-microgrid-firm-utilityinnovation-group/（轉述公告；10-Q 僅列合併協議）"),
        ("[Assumed] 資料中心設施產品占營收比例（dcp）：公司未揭露，依 10-Q 營收分類推估", "本模型假設")), 3):
    SO.cell(i, 1, a); SO.cell(i, 2, b)
SO.column_dimensions["A"].width = 120; SO.column_dimensions["B"].width = 110

exec(open(os.path.join(HERE, "build_part3.py"), encoding="utf-8").read())
