# AMD v0.4.5／v0.4.6（2026-11-10）後處理：AMD v0.4.4 驗收報告的採用項。
# 採用規劃書：Drive 03-AI半導體/Claude outputs/20261110_AMD驗收v0.4.4_採用規劃書.md（Andy 2026-11-10 07:52 確認「照預設」：J1(a)、J2(b)、J3(b)、J4(a)）。
# build_amd.py 在存檔前呼叫 apply(wb, here, ctx)。ctx["V46"]＝False 時只做展示與治理（數值不變，v0.4.5）；True 時加做數值項（v0.4.6）。
#   v0.4.5：A0 原則 1、2、3、5、10、13、14、15、20、32、36（build_id／狀態、單一標記與來源誘因、Hist 頁、as_of、反解、DCF 單法欄、邊界揭露）；
#           A3 三個日期分列；A4 舊法移出主敘述；A5 倍數標分母；A6 手寫舊數字改公式；A8 風險登錄與可證偽指標；A9 命題 1 條件式；
#           B4 non-GAAP → FCFF 調節橋與現金稅率；B6 β 依據；B7 承重外部輸入的定位
#   v0.4.6：A1 認股權證閘門式（build_amd.py）；J2(b) 情境 A／C 帶顆數與 CPU 路徑（build_amd.py）；原則 26 目標時點、原則 11 評價股數、
#           原則 35 股票薪酬比率最近四季；目標價 Sens 業務主因（整本重算快照 amd/sens.json：A2、B3、B4、B5、J1 0％／50％）；Recon 世代售價係數反解
import hashlib, json, os, re
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

BOLD = Font(bold=True); BLUE = Font(color="0000FF"); HDR = PatternFill("solid", fgColor="DDEBF7")
AS_OF, CUTOFF, BUILT, VAL_D = "2026-11-05", "2026-11-05", "2026-11-10", "2026-12-31"
E8K = "https://www.sec.gov/Archives/edgar/data/2488/"
PR3 = E8K + "000000248826000188/q32026991.htm"; PR4 = E8K + "000000248826000014/q42025991.htm"
Q10 = E8K + "000000248826000191/amd-20260926.htm"
CF25 = "AMD 8-K 2026-02-03（FY2025 現金流量表）"; CF26 = "AMD 8-K 2026-11-03（2026 前三季現金流量表）"
XB = "SEC XBRL companyfacts（CIK 0000002488）"


def lit(t):   # Excel 公式字串常數上限 255 字元：切段以 & 相接
    t = t.replace('"', '""')
    return "&".join('"' + t[i:i + 200] + '"' for i in range(0, len(t), 200)) or '""'


def lab(ws, prefix, start=1, col=1):
    for r in range(start, ws.max_row + 1):
        v = ws.cell(r, col).value
        if isinstance(v, str) and v.strip().startswith(prefix): return r
    raise KeyError((ws.title, prefix))


def inrow(wb, key):
    IN = wb["In"]
    return next(r for r in range(1, IN.max_row + 1) if IN.cell(r, 1).value == key)


def dn(wb, name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


def in_set(wb, key, col, val):
    wb["In"].cell(inrow(wb, key), col).value = val


# ---------- 原則 14：Hist 頁 ----------
HIST_Y = [  # (鍵, 項目, FY2023, FY2024, FY2025, 2026 前三季, 標記, 來源)
    ("rev", "營收", 22.68, 25.785, 34.639, 35.166, "Verified", XB + "：Revenue；2026 前三季 " + CF26),
    ("da", "折舊攤銷（不含併購無形資產攤銷）", None, 0.671, 0.75, 0.666, "Verified", CF25 + "、" + CF26),
    ("cx", "資本支出（持續營業）", None, 0.636, 0.974, 1.845, "Verified", CF25 + "、" + CF26),
    ("sbc", "股票薪酬", None, 1.407, 1.638, 1.513, "Verified", CF25 + "、" + CF26 + "；" + PR3),
    ("ar", "應收帳款（年底）", 4.323, 6.192, 6.315, None, "Verified", XB),
    ("inv", "存貨（年底）", 4.351, 5.734, 7.92, None, "Verified", XB),
    ("pre", "預付及其他流動資產（年底）", 2.312, 1.991, 2.16, None, "Verified", XB),
    ("ap", "應付帳款（年底）", 2.055, 2.466, 2.929, None, "Verified", XB),
    ("acc", "應計負債（年底）", 3.082, 4.26, 5.25, None, "Verified", XB),
]
HIST_P = [  # (鍵, 項目, 值, 標記, 來源)
    ("sbc_q425", "4Q25 股票薪酬", 0.486, "Verified", "公司 4Q25 新聞稿 non-GAAP 調節 " + PR4),
    ("rev_q425", "4Q25 營收", 10.27, "Verified", "同上"),
    ("shout", "流通股數（2026-10-28，十億股）", 1.640087251, "Verified", "AMD 10-Q（季度結束 2026-09-26）封面：1,640,087,251 股 " + Q10),
    ("dil_q3", "3Q26 GAAP 稀釋加權股數（十億股）", 1.661, "Verified", "公司 3Q26 新聞稿 " + PR3),
    ("bas_q3", "3Q26 GAAP 基本加權股數（十億股）", 1.636, "Verified", "同上"),
    ("tsm", "RSU 與選擇權的庫藏股法稀釋（十億股）＝稀釋 − 基本加權股數（代理）", "=h_dil_q3-h_bas_q3", "Derived",
     "10-Q 未揭露期末未既得 RSU 與選擇權數；以 3Q26 GAAP 稀釋與基本加權股數之差代理庫藏股法稀釋"),
    ("taxpaid_9m", "2026 前三季已付所得稅（扣退稅）", 0.833, "Verified", "AMD 10-Q 2026-09-26 現金流量表補充揭露 " + Q10),
    ("gpti_9m", "2026 前三季 GAAP 稅前利益（持續營業）", 6.023, "Verified", "同上（損益表）"),
    ("amort_9m", "2026 前三季併購無形資產攤銷", 1.629, "Verified", "同上（現金流量表；銷貨成本 778＋營業費用 851，US$M）"),
    ("acq_9m", "2026 前三季收購相關及其他費用（non-GAAP 排除）", 0.093, "Verified", "公司 3Q26 新聞稿 non-GAAP 調節 " + PR3),
    ("legal_9m", "2026 前三季法律或有損失（non-GAAP 排除）", 0.024, "Verified", "同上"),
    ("wfv_q3", "3Q26 認股權證負債公允價值變動（non-GAAP 排除）", 1.097, "Verified", "同上"),
    ("w_unrec", "未歸屬認股權證之授予日公允價值（未來認列為營收減項）", 4.5, "Verified", "AMD 10-Q 2026-09-26 附註：unrecognized grant-date fair value $4.5 billion，would be recorded as a reduction of revenue " + Q10),
    ("w_gfv", "認股權證加權平均授予日公允價值（美元／股）", 14.25, "Verified", "同上"),
]


def hist_sheet(wb):
    H = wb.create_sheet("Hist", index=wb.sheetnames.index("Q") + 1)
    H["A1"] = ("Hist — 歷史財務與評價用資料（v0.4.5 起；驗收原則 14）：In 頁的折舊、資本支出、營運資金、股票薪酬、評價股數與 DCF 調節橋只引用本頁儲存格"
               "（定義名稱 h_*），不在公式內寫常數。金額：十億美元；股數：十億股")
    for j, h in enumerate(["鍵", "項目", "FY2023", "FY2024", "FY2025", "2026 前三季", "標記", "來源"], 1):
        c = H.cell(3, j, h); c.font = BOLD; c.fill = HDR
    r = 4
    for k, labl, *vals, tag, src in HIST_Y:
        H.cell(r, 1, f"h_{k}"); H.cell(r, 2, labl)
        for j, (y, v) in enumerate(zip(("23", "24", "25", "9m26"), vals)):
            if v is None: continue
            c = H.cell(r, 3 + j, v); c.font = BLUE; c.number_format = "#,##0.000"; dn(wb, f"h_{k}_{y}", f"Hist!${'CDEF'[j]}${r}")
        H.cell(r, 7, tag); H.cell(r, 8, src); r += 1
    r += 1
    for j, h in enumerate(["鍵", "項目", "值", "", "", "", "標記", "來源"], 1):
        c = H.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    for k, labl, v, tag, src in HIST_P:
        H.cell(r, 1, f"h_{k}"); H.cell(r, 2, labl); c = H.cell(r, 3, v); c.number_format = "#,##0.000000"
        if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
        dn(wb, f"h_{k}", f"Hist!$C${r}"); H.cell(r, 7, tag); H.cell(r, 8, src); r += 1
    r += 1
    for k, labl, f in (("asof", "資料基準日（現價日）", "=DATE(2026,11,5)"), ("cutoff", "資料截止日（最晚一筆資料：股價與同業報價）", "=DATE(2026,11,5)"),
                       ("built", "成品日（本版建置）", "=DATE(2026,11,10)"), ("valdate", "DCF 評價日", "=DATE(2026,12,31)"), ("bsdate", "資產負債表日", "=DATE(2026,9,26)")):
        H.cell(r, 1, f"h_{k}"); H.cell(r, 2, labl); H.cell(r, 3, f).number_format = "yyyy-mm-dd"; dn(wb, f"h_{k}", f"Hist!$C${r}"); H.cell(r, 7, "—"); r += 1
    H.column_dimensions["A"].width = 16; H.column_dimensions["B"].width = 56
    for c in "CDEF": H.column_dimensions[c].width = 13
    H.column_dimensions["G"].width = 12; H.column_dimensions["H"].width = 100
    return H


def hist_links(wb, V46P, OVR):
    """原則 14：In 頁公式內的歷史常數改引用 Hist（v0.4.5 數值不變）；v0.4.6 'sbc' 改最近四季（原則 35）"""
    S = {
        ("fc_da", 5): "=(h_da_24/h_rev_24+h_da_25/h_rev_25+h_da_9m26/h_rev_9m26)/3",
        ("fc_cx", 5): "=(h_cx_24/h_rev_24+h_cx_25/h_rev_25+h_cx_9m26/h_rev_9m26)/3",
        ("fc_wc", 5): "=" + "+".join(f"(h_ar_{y}+h_inv_{y}+h_pre_{y}-h_ap_{y}-h_acc_{y})/h_rev_{y}" for y in ("23", "24", "25")).join(("(", ")/3")),
        ("sbc_rev", 5): ("=(h_sbc_q425+h_sbc_9m26)/(h_rev_q425+h_rev_9m26)" if "sbc" in V46P else
                         "=(h_sbc_24/h_rev_24+h_sbc_25/h_rev_25+h_sbc_9m26/h_rev_9m26)/3"),
    }
    for (k, c), f in S.items():
        cell = wb["In"].cell(inrow(wb, k), c)
        if k in OVR: continue
        assert isinstance(cell.value, str) and cell.value.startswith("="), (k, c, cell.value)
        cell.value = f
    if "sbc" in V46P:
        r = inrow(wb, "sbc_rev"); IN = wb["In"]
        IN.cell(r, 4, "=h_sbc_9m26/h_rev_9m26"); IN.cell(r, 6, "=(h_sbc_24/h_rev_24+h_sbc_25/h_rev_25+h_sbc_9m26/h_rev_9m26)/3")
        dn(wb, "sbc_rev_lo", f"In!$D${r}"); dn(wb, "sbc_rev_hi", f"In!$F${r}")
        IN.cell(r, 8, "股票薪酬 ÷ 營收，最近四季（4Q25–3Q26）：(486＋1,513) ÷ (10,270＋35,166)（US$M；公司 4Q25 新聞稿、" + CF26 + "；Hist 頁）")
        IN.cell(r, 9, "原則 35（v0.4.6）：比率以最近四季實績為起點、全期持平；低＝2026 前三季、高＝FY2024–2026 前三季三期平均（v0.4.5 以前的基準）。客戶認股權證的非現金影響不另扣，以股數稀釋計（w_val）")


# ---------- v0.4.6 數值項（build_amd.py 以外者）----------
def v46_values(wb, V46P, OVR):
    IN = wb["In"]
    if "tt" in V46P:   # 原則 26：目標時點＝現價日＋12 個月（評價日 2026-12-31 不變；不需驗收清單 X20）
        rr = inrow(wb, "t_tgt")
        IN.cell(rr, 2, "目標時點距評價日（目標時點＝現價日＋12 個月）"); IN.cell(rr, 5, "=(EDATE(h_asof,12)-h_valdate)/365")
        IN.cell(rr, 8, "目標價為現價日起 12 個月的價格，與 Comparable 的 12 個月後前瞻 EPS 同一時點（原則 26）"); IN.cell(rr, 9, "v0.4.5 以前＝1 年（評價日＋12 個月，即現價日＋約 14 個月）")
    if "sh" in V46P:   # 原則 11：評價日股數
        rr = inrow(wb, "sh_v")
        IN.cell(rr, 8, "流通股 1,640.09 百萬股（10-Q 封面 2026-10-28）＋RSU 與選擇權庫藏股法稀釋（3Q26 稀釋 − 基本加權股數 25 百萬股，代理）＋評價用客戶認股權證股數"
                        + ("（w_val：到期日前歸屬 GW × wf；原則 37）" if "w" in V46P else "（上限 w_tot 全額）"))
        IN.cell(rr, 9, "個股DCF標準規格第 3 節（原則 11）；v0.4.5 以前＝3Q26 稀釋加權股數 1.661＋認股權證全額 0.32")


def q_overrides(wb, OVR):
    """敏感度快照用：GPU 錨點占比（J3）、Q4 其他分部誤差（A2）"""
    Q = wb["Q"]
    if "_gsh" in OVR:   # FY2026 各季 GPU ＝ 占比 × 資料中心；非 GPU 為餘額（k_rec 隨之重算）
        x = OVR["_gsh"]
        for c in "GHIJ":
            Q[f"{c}6"] = f"={c}5*{x}"; Q[f"{c}7"] = f"={c}5-{c}6"
    if "_oseg" in OVR:  # Q4 其他分部合計誤差全數由資料中心吸收（Client 加 d；資料中心＝指引 − 其他分部）
        Q["J8"] = f"=I8*(1+q4_cli)+({OVR['_oseg']})"
    if "_rmhi" in OVR:
        V = wb["Val"]; V.cell(lab(V, "夾值上限"), 2).value = OVR["_rmhi"]


# ---------- 原則 2：單一主要標記與來源誘因 ----------
def incentive(src):
    t = str(src or "")
    if any(k in t for k in ("共識", "賣方", "分析師", "Yahoo", "stockanalysis", "StockAnalysis", "tickflow", "HSBC", "S&P Global", "Motley", "ts2.tech")):
        return "賣方分析師或資料商：評等與交易業務誘因，目標價與預估傾向跟隨股價"
    if any(k in t for k in ("法說", "逐字稿", "MarketBeat", "platformaeronaut", "二手")):
        return "公司管理層口頭陳述（二手轉述）：股價與市場預期管理誘因，未經查核、可能偏樂觀"
    if any(k in t for k in ("OpenAI", "Oracle", "Meta", "Anthropic")):
        return "交易對手與公司共同發布：雙方皆有宣傳合作規模的誘因；GW 為上限與意向，非已確認訂單"
    if any(k in t for k in ("指引", "公司", "新聞稿", "8-K", "UNITS", "market.py")):
        return "公司指引、自揭數字或上游模型轉載的公司／產業說法：預期管理誘因，口徑由提供者定義"
    return "利害關係方提供：未經獨立查核"


def norm_tags(wb):
    n = 0
    for sh, tcol, scol, icol in (("In", 7, 8, 10), ("Q", 12, 13, 15)):
        ws = wb[sh]
        hr = next(r for r in range(1, 15) if ws.cell(r, tcol).value == "標記")
        c = ws.cell(hr, icol, "來源誘因"); c.font = BOLD; c.fill = HDR; ws.column_dimensions[c.column_letter].width = 50
        for r in range(hr + 1, ws.max_row + 1):
            t = ws.cell(r, tcol).value
            if not isinstance(t, str) or t == "標記": continue
            if "／" in t:
                parts = [p.strip() for p in re.split(r"／|（", t)]
                prim = "Interested-party" if any(p.startswith("Interested-party") for p in parts) else ("Analogy" if "Analogy" in parts[0] else parts[0])
                if prim.startswith("Verified") and "Interested-party" in t: prim = "Interested-party"
                ws.cell(r, tcol).value = prim
                note = f"欄別標記（v0.4.4 混標「{t}」）：主要標記取最近一期原文；實績季為財報（Verified）、Q4 2026 為公司指引（Interested-party）" if sh == "Q" else f"v0.4.4 混標「{t}」：基準 1.0 為假設（Assumed 性質），高值為容量比類比（Analogy）；主要標記取 Analogy（承重值有類比依據）"
                ws.cell(r, icol, (incentive("指引") + "；" + note) if prim == "Interested-party" else note); n += 1; t = prim
            if sh == "In" and t == "Verified":
                base = ws.cell(r, 5).value
                if isinstance(base, str) and base.startswith("=") and re.search(r"[+\-*/]", base[1:]) and not base.startswith("=DATE("):
                    ws.cell(r, tcol).value = t = "Derived"; n += 1
            if t not in ("Verified", "Interested-party", "Analogy", "Assumed", "Derived", "—"):
                ws.cell(r, tcol).value = t = "Verified" if t.startswith("Verified") else ("Assumed" if "Assumed" in t else ("Derived" if "Derived" in t else "Interested-party")); n += 1
            if t == "Interested-party" and not ws.cell(r, icol).value:
                ws.cell(r, icol, incentive(f"{ws.cell(r, scol).value or ''}｜{ws.cell(r, 2).value or ''}"))
    return n


# ---------- 原則 1：build_id、輸入指紋、模型狀態 ----------
def build_id(here, stage):
    h = hashlib.sha256(stage.encode())
    for fn in sorted(["build_amd.py", "v045_post.py", "px1105_post.py", "tp_v2.py", "sens.json", "segments.json", "hist_frozen.json"]):
        p = os.path.join(here, fn)
        if os.path.exists(p): h.update(fn.encode()); h.update(open(p, "rb").read())
    return f"amd-v{stage}-{h.hexdigest()[:10]}"


def governance(wb, SJ, bid, OVR):
    SE = wb["Sens"]; IN = wb["In"]
    hr = next(r for r in range(1, 15) if IN.cell(r, 7).value == "標記")
    last = next(r for r in range(hr + 1, IN.max_row + 1) if str(IN.cell(r, 1).value or "").startswith("情境驅動對照")) - 1
    rng = f"In!$D${hr + 1}:$F${last}"
    r = SE.max_row + 3; G = {}
    SE.cell(r, 1, "模型治理（v0.4.5 起；原則 1）：Sens 主表與 EPS 案例為活頁簿公式；「業務主因（整本重算）」表與 HTML 是建置快照。只在 Excel 改 In 頁而未重跑 python3 amd/sens_amd.py 與建置時，下方模型狀態顯示 STALE").font = BOLD; r += 1
    SE.cell(r, 1, "輸入指紋（本檔即時重算；In 頁低／基準／高欄，情境驅動對照以上）"); SE.cell(r, 4, f"=SUMPRODUCT(N(+{rng}),ROW({rng})*3+COLUMN({rng}))").number_format = "0.000000"; G["fp"] = r; r += 1
    base = (SJ or {}).get("_base") if not OVR else None
    SE.cell(r, 1, "輸入指紋（建置快照的基準檔）")
    if base: SE.cell(r, 4, base["fp"]).number_format = "0.000000"
    fb = r; r += 1
    SE.cell(r, 1, "加權目標價（建置快照的基準檔）")
    if base: SE.cell(r, 4, base["tp"]).number_format = "#,##0.00"
    tb = r; r += 1
    tpc = f"Val!$B${lab(wb['Val'], '目標價（DCF 與 Comparable 加權）')}"
    SE.cell(r, 1, "模型狀態（CURRENT＝Sens 快照、HTML、README 與本檔輸入一致；STALE＝In 頁已改、須重跑 python3 amd/sens_amd.py 與建置）").font = BOLD
    SE.cell(r, 4, f'=IF(AND(ISNUMBER(D{fb}),ISNUMBER(D{tb}),ABS(D{G["fp"]}-D{fb})<=0.000000001*MAX(1,ABS(D{fb})),ABS({tpc}-D{tb})<0.005),"CURRENT","STALE")'); G["st"] = r; r += 1
    SE.cell(r, 1, "build_id（README、Sens、HTML 三處相同）"); SE.cell(r, 4, bid); G["bid"] = r
    return G


# ---------- 原則 32：目標價 Sens 加 DCF 單法美元欄；業務主因（整本重算快照）----------
BIZ = [("AMD 顆數係數 ×0.8／×1.2（三情境同乘；認股權證歸屬隨之）", "uf", "In 頁 uf 基準換成低／高（大單延後／提前）"),
       ("世代售價係數 高（隨 HBM 容量；三情境）", "pg", "FY2027 ×1.25、FY2028 起 ×1.5（只有高；成本隨營收等比例，見 Sens 讀法）"),
       ("HBM 漲價吸收比例 0／1（三情境）", "abs", "全額轉嫁／全額吸收"),
       ("GPU 毛利率 40%／55%", "gm_gpu", "非 GPU 毛利率在 FY2026 重新校準"),
       ("資料中心 CPU 成長 低／高（三情境同用）", "cpu", "低 +30／+10／+5%、高（公司說法）+70／+40／+25%"),
       ("FY2026 GPU 占資料中心 40%／60%（J3；整本重算 k_rec）", "gsh", "基準 53%；各季 GPU＝占比 × 資料中心，非 GPU 為餘額；k_rec、非 GPU 毛利率隨之重算（原則 39）"),
       ("Q4 其他分部合計 −1.0／+1.0（US$B；A2）", "oseg1", "誤差全數由 Q4 資料中心 GPU 吸收（餘額法），整本重算 k_rec"),
       ("Q4 其他分部合計 −0.5／+0.5（US$B；A2）", "oseg05", "同上"),
       ("認股權證歸屬 0％／50％（股價門檻與條件；J1）", "wf", "wf＝0／0.5（基準 1）"),
       ("Comparable 調整係數不夾值（2.27；J4）", "rmhi", "夾值上限 2.0 → 不夾值（只有高）"),
       ("non-GAAP 稅率改現金稅率（B4）", "tx", "tx_f 換成 2026 前三季已付所得稅 ÷ non-GAAP 稅前利益（只有低）")]


def sens_tp(wb, SJ, V46):
    SE = wb["Sens"]
    r0 = lab(SE, "目標價參數敏感度"); hr = r0 + 1
    SE.cell(hr, 9, "低：DCF 單法變化（美元）").font = BOLD; SE.cell(hr, 10, "高：DCF 單法變化（美元）").font = BOLD
    r = hr + 1
    while SE.cell(r, 1).value and not str(SE.cell(r, 1).value).startswith("檢查"):
        is_w = SE.cell(r, 8).value == "tp_w_dcf"
        for cv, cd in (("C", "I"), ("E", "J")):
            SE[f"{cd}{r}"] = '="—（權重案例）"' if is_w else f'=IF(ISNUMBER({cv}{r}),({cv}{r}-G{r})/tp_w_dcf,"—")'
            SE[f"{cd}{r}"].number_format = "+#,##0.00;-#,##0.00"
        r += 1
    SE.cell(r0, 1).value = str(SE.cell(r0, 1).value) + "（v0.4.5 起 I、J 欄＝DCF 單法目標價變化，美元；Comparable 不變的案例＝加權變化 ÷ DCF 權重）"
    G = {"tp0": r0, "tp_end": r - 1}
    if not V46: return G
    r = SE.max_row + 3
    SE.cell(r, 1, "目標價敏感度：業務主因（整本重算快照；v0.4.6 起，驗收 A2、B3、B4、B5、J1；原則 32、39）").font = BOLD; r += 1
    SE.cell(r, 1, "每案把輸入換掉後整本重建、LibreOffice 重算（amd/sens_amd.py → amd/sens.json），讀加權目標價與 DCF 目標價；不是活頁簿公式，In 頁改過須重跑（模型狀態見本頁底部）"); r += 1
    for j, h in enumerate(["案例", "設定", "低：加權目標價", "低：相對基準", "高：加權目標價", "高：相對基準", "基準加權目標價", "鍵", "低：DCF 變化（美元）", "高：DCF 變化（美元）"], 1):
        c = SE.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1; G["biz0"] = r; G["bizrow"] = {}
    b = (SJ or {}).get("_base", {})
    for name, key, setting in BIZ:
        d = (SJ or {}).get("biz", {}).get(key, {})
        SE.cell(r, 1, name); SE.cell(r, 2, setting); SE.cell(r, 8, key); G["bizrow"][key] = r
        for side, ct, cr, cdl in (("低", 3, 4, 9), ("高", 5, 6, 10)):
            v = d.get(side)
            if v and b:
                SE.cell(r, ct, v["tp"]).number_format = "#,##0.00"; SE.cell(r, cr, v["tp"] / b["tp"] - 1).number_format = "+0.0%;-0.0%"
                SE.cell(r, cdl, v["dcf"] - b["dcf"]).number_format = "+#,##0.00;-#,##0.00"
            else:
                SE.cell(r, ct, "—")
        if b: SE.cell(r, 7, b["tp"]).number_format = "#,##0.00"
        r += 1
    G["biz1"] = r - 1
    return G


# ---------- Recon：同口徑註記、反解 ----------
def recon(wb, SJ, V46):
    RC = wb["Recon"]; G = {}
    r = lab(RC, "FY2027 資料中心非 GPU 年增 vs 公司伺服器 CPU")
    RC.cell(r, 10, str(RC.cell(r, 10).value or "") + "（原則 33：模型的非 GPU 含 DPU、FPGA 等，公司說法只指伺服器 CPU；口徑不同，差距只作方向參考）")
    r = RC.max_row + 2
    RC.cell(r, 1, "E. 反解：支持現價與評等翻轉所需的值（v0.4.5 起；原則 4、36）").font = BOLD; r += 1
    V = wb["Val"]
    dcf = f"Val!$B${lab(V, 'DCF 目標價（情境加權）')}"
    feps = f"Val!$B${lab(V, '12 個月後的前瞻 EPS（情境加權）')}"; pmed = f"Val!$B${lab(V, '同業前瞻 P/E 中位數')}"
    up = f"Val!$B${lab(V, '上檔空間：加權目標價')}"
    RC.cell(r, 1, "評等級距邊界：升一級（上檔空間）"); RC.cell(r, 2, f"=IF({up}<-0.5,-0.5,IF({up}<-0.25,-0.25,IF({up}<0.25,0.25,IF({up}<0.5,0.5,NA()))))").number_format = "+0%;-0%"
    RC.cell(r, 10, "級距：大於 +50% 強力買進、+25%～+50% 買進、−25%～+25% 中立、−50%～−25% 放空、小於 −50% 強力放空（Andy 2026-10-25；與目標價總表同一套邊界，原則 24）"); G["e_up"] = r; r += 1
    RC.cell(r, 1, "評等級距邊界：降一級（上檔空間）"); RC.cell(r, 2, f"=IF({up}>0.5,0.5,IF({up}>0.25,0.25,IF({up}>-0.25,-0.25,IF({up}>-0.5,-0.5,NA()))))").number_format = "+0%;-0%"; G["e_dn"] = r; r += 1
    for j, h in enumerate(["項目", "需要的值", "模型現值", "需要的變動"], 1): RC.cell(r, j, h).font = BOLD
    r += 1
    w = "tp_w_dcf"; G["inv"] = {}
    tgt = {"px": "px", "up": f"px*(1+B{G['e_up']})", "dn": f"px*(1+B{G['e_dn']})"}
    nm = {"px": "支持現價", "up": "評等升一級", "dn": "評等降一級"}
    for t in ("up", "dn", "px"):
        for kind, f, fmt, cur in (("pe", f"=({tgt[t]}-{w}*{dcf})/((1-{w})*{feps})", "0.0", f"={pmed}*rm_factor"),
                                  ("eps", f"=({tgt[t]}-{w}*{dcf})/((1-{w})*{pmed}*rm_factor)", "0.00", f"={feps}"),
                                  ("dcf", f"=({tgt[t]}-(1-{w})*tp_comp_adj)/{w}", "0.00", f"={dcf}")):
            labl = {"pe": f"{nm[t]}所需的 Comparable 採用倍數（同業中位數 × 相對倍數係數；DCF 不變）",
                    "eps": f"{nm[t]}所需的 12 個月後前瞻 EPS（同業中位數 × 係數；DCF 不變）",
                    "dcf": f"{nm[t]}所需的 DCF 目標價（Comparable 不變）"}[kind]
            RC.cell(r, 1, labl); RC.cell(r, 2, f).number_format = fmt; RC.cell(r, 3, cur).number_format = fmt
            RC.cell(r, 4, f"=B{r}/C{r}-1").number_format = "+0.0%;-0.0%"
            RC.cell(r, 10, "閉式：加權目標價＝w × DCF＋（1 − w）× 倍數 × EPS × 係數；另一法固定。反解值是目標價等於該點的值")
            G["inv"][f"{kind}_{t}"] = r; r += 1
    if V46 and SJ and SJ.get("pk"):
        pk = SJ["pk"]
        for t in ("up", "dn", "px"):
            v = pk.get(t)
            RC.cell(r, 1, f"{nm[t]}所需的世代售價係數乘數（FY2027 起三情境同乘；整本重算內插）"); RC.cell(r, 2, v if v is not None else "超出測試範圍").number_format = "0.00"; RC.cell(r, 3, 1).number_format = "0.00"
            if v is not None: RC.cell(r, 4, v - 1).number_format = "+0%;-0%"
            RC.cell(r, 10, "amd/sens_amd.py：乘數 " + "、".join(f"{p['k']:g}" for p in pk["pts"]) + " 各整本重算，線性內插；售價與成本同比例（GPU 毛利率不變）")
            G["inv"]["pk_" + t] = r; r += 1
    return G


# ---------- Val：舊法標示、期間並列、說明 ----------
def val(wb, ctx):
    V = wb["Val"]; j2 = "j2" in ctx["V46P"]
    V["A1"].value = ("Val — 估值（美元；non-GAAP EPS）。目標價＝DCF（標準版 FCFF＋WACC，第 0.1 節）與 Comparable（同業前瞻 P/E 中位數 × 相對倍數調整 × 12 個月後前瞻 EPS，第 0.2、0.5 節）加權（第 0.3 節）。"
                     "下方「舊法：規則值情境加權」與十年 P/E 各節是舊法，只作參考，不進結論（v0.4.5 起；驗收 A4、原則 42）")
    for rr in range(1, V.max_row + 1):
        a = V.cell(rr, 1).value
        if isinstance(a, str) and a.startswith(("舊法：規則值情境加權", "P/E：十年歷史倍數")) and "不進結論" not in a:
            V.cell(rr, 1).value = a + "【舊法，不進結論】"
        e = V.cell(rr, 5).value
        if j2 and isinstance(e, str) and e.startswith("FY2027 EPS 三情境相同"):
            V.cell(rr, 5).value = "FY2027 EPS 三情境不同（v0.4.6 起情境 A／C 帶 AMD 顆數係數與資料中心 CPU 成長路徑；驗收 J2）"
    rc = lab(V, "投資建議")
    V.cell(rc, 5).value = str(V.cell(rc, 5).value) + "。本列為估值評等，未評估借券、流通量與時點"
    for lbl in ("現價 ÷ 共識 FY2027 EPS",):
        rr = lab(V, lbl); V.cell(rr, 5).value = "分母：B＝stockanalysis 共識 FY2027 EPS 15.76、C＝tickflow 四季合計 11.23（皆 2026-10-23，3Q26 財報前；FY 口徑，非 NTM）；兩來源口徑與日期未釐清前不據以判斷市場錯估（原則 43）"
    rr = lab(V, "現價 ÷ 共識 FY2028 EPS"); V.cell(rr, 5).value = "分母：Motley Fool 經 ts2.tech 轉述約 20 美元（2026-07-30，二手；FY 口徑）"
    for lbl in ("現價 ÷ FY2026 EPS", "現價 ÷ FY2027 EPS", "現價 ÷ FY2028 EPS", "現價 ÷ FY2029 EPS"):
        rr = lab(V, lbl)
        if not V.cell(rr, 5).value or j2: V.cell(rr, 5).value = "分母：本模型各情境 non-GAAP EPS（FY 口徑；B／C／D 欄＝情境 A／B／C）；現價日 2026-11-05"
    # 原則 18 延伸（原則 38）：夾值揭露
    rr = lab(V, "相對倍數調整係數")
    V.cell(rr, 3).value = f'=IF(B{lab(V, "原始比值")}>B{lab(V, "夾值上限")},"已達夾值上限：原始比值 "&TEXT(B{lab(V, "原始比值")},"0.00")&"（承重輸入；Sens 業務主因表有不夾值案例）","")'
    G = {}
    r = V.max_row + 2
    V.cell(r, 1, "0.6 期間並列（v0.4.5 起；原則 5）").font = BOLD; r += 1
    V.cell(r, 1, "折現期間：DCF 評價日 → 目標時點（年）"); V.cell(r, 2, "=t_tgt").number_format = "0.00"; V.cell(r, 5, '="評價日 "&TEXT(h_valdate,"yyyy-mm-dd")&"；DCF 每股價值以股權成本推 t_tgt 年"'); G["per1"] = r; r += 1
    V.cell(r, 1, "持有期間：現價日 → 目標時點（年）"); V.cell(r, 2, "=(h_valdate+t_tgt*365-h_asof)/365").number_format = "0.00"
    V.cell(r, 5, '="現價日 "&TEXT(h_asof,"yyyy-mm-dd")&"；上檔空間是這段期間的價格變化，不是年化報酬；Comparable 的前瞻 EPS 對應現價日＋約 12 個月"'); G["per2"] = r
    return G


# ---------- B4：non-GAAP → FCFF 調節橋（DCF 頁）----------
def dcf_bridge(wb):
    D = wb["DCF"]; r = D.max_row + 3
    D.cell(r, 1, "non-GAAP → FCFF 調節橋（v0.4.5 起；驗收 B4、原則 40）：公司 non-GAAP 調節各項在本 DCF 的處理").font = BOLD; r += 1
    for j, h in enumerate(["項目", "2026 前三季（US$B）", "現金／非現金", "DCF 處理"], 1):
        c = D.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1; r0 = r
    rows = [("股票薪酬", "=h_sbc_9m26", "非現金（以股票支付）", "EBIT 扣除（sbc_rev × 營收；股權報酬視為費用，個股DCF標準規格第 1 節）"),
            ("併購無形資產攤銷", "=h_amort_9m", "非現金", "不計；折舊攤銷 fc_da 也不含此攤銷（與 non-GAAP EBIT 同口徑）"),
            ("收購相關及其他費用", "=h_acq_9m", "現金", '="未扣（占營收 "&TEXT(h_acq_9m/h_rev_9m26,"0.00%")&"；全數扣除對 DCF 影響小於 1%，不另建 Sens）"'),
            ("法律或有損失", "=h_legal_9m", "現金（或有）", '="未扣（占營收 "&TEXT(h_legal_9m/h_rev_9m26,"0.00%")&"）"'),
            ("認股權證負債公允價值變動（3Q26）", "=h_wfv_q3", "非現金", "不計；經濟成本以股數稀釋計（評價股數 sh_v 含 w_val，原則 37）"),
            ("認股權證授予日公允價值之營收減項（未認列餘額）", "=h_w_unrec", "非現金", "公司報表營收已扣除（GAAP＝non-GAAP 營收），模型營收由報表校準，故已內含；與股數稀釋部分重複，金額約占大單營收 1–2%，不調整"),
            ("長期投資損益", "—", "非營業", "不計；權益投資以帳面加入淨現金（inv_v）")]
    for a, b, c, d in rows:
        D.cell(r, 1, a); x = D.cell(r, 2, b); x.number_format = "#,##0.000"; D.cell(r, 3, c); D.cell(r, 4, d); r += 1
    D.cell(r, 1, "現金稅率＝已付所得稅 ÷ non-GAAP 稅前利益（2026 前三季）"); D.cell(r, 2, "=h_taxpaid_9m/(SUM(Q!G15:I15)+SUM(Q!G20:I20))").number_format = "0.0%"
    dn(wb, "cash_tx", f"DCF!$B${r}")
    D.cell(r, 3, '="模型稅率 tx_f "&TEXT(tx_f,"0.0%")&"；差 "&TEXT((tx_f-B' + str(r) + ')*100,"0.0")&" 個百分點"')
    D.cell(r, 4, '=IF(ABS(tx_f-B' + str(r) + ')>0.03,"差距大於 3 個百分點：Sens 業務主因表有「稅率改現金稅率」案例（原則 40）","差距在 3 個百分點內")'); r += 1
    D.cell(r, 1, "對照：已付所得稅 ÷ GAAP 稅前利益（2026 前三季；GAAP 稅前含認股權證公允價值損失等非現金項）"); D.cell(r, 2, "=h_taxpaid_9m/h_gpti_9m").number_format = "0.0%"
    D.column_dimensions["D"].width = max(D.column_dimensions["D"].width or 0, 13)
    return r0


# ---------- README ----------
def readme(wb, ctx, bid, G):
    RD = wb["README"]; V = wb["Val"]; VER = "v" + ctx["STAGE"]; P46 = ctx["V46P"]; j2 = "j2" in P46
    tp = f"Val!B{lab(V, '目標價（DCF 與 Comparable 加權）')}"; up = f"Val!B{lab(V, '上檔空間：加權目標價')}"
    rec = f"Val!B{lab(V, '投資建議')}"; dcf = f"Val!B{lab(V, 'DCF 目標價（情境加權）')}"; cmp_ = "tp_comp_adj"
    pmed = f"Val!B{lab(V, '同業前瞻 P/E 中位數')}"; band = lab(V, "價格帶（DCF")
    feps = f"Val!B{lab(V, '12 個月後的前瞻 EPS（情境加權）')}"
    SUMR = ctx["SUMR"]; eps = lambda s, c: f"PL!{c}{SUMR[(s, 'eps')]}"
    # A3：首列三個日期＋build_id（公式）
    a1 = str(RD["A1"].value)
    head, _, rest = a1.partition("　2026-11-04；")
    RD["A1"].value = ("=" + lit(head.replace("v0.4.4", VER)) + '&"　資料截止日 "&TEXT(h_cutoff,"yyyy-mm-dd")&"；成品日 "&TEXT(h_built,"yyyy-mm-dd")&"；股價基準日 "&TEXT(h_asof,"yyyy-mm-dd")&"（v0.4.5 起三個日期分列，驗收 A3）；build_id '
                      + bid + '。版本：' + '"&' + lit(rest))
    for r in range(1, RD.max_row + 1):
        v = RD.cell(r, 1).value
        if not isinstance(v, str): continue
        w = v
        if v.startswith("模型命題：") and "股價約為 FY2027 共識 EPS 的 33 倍" in v:   # A6
            a, b = v.split("股價約為 FY2027 共識 EPS 的 33 倍", 1)
            w = "=" + lit(a) + '&"股價約為 FY2027 共識 EPS 的 "&TEXT(px/cons_eps27,"0")&" 倍（分母：stockanalysis 共識 15.76，2026-10-23，3Q26 財報前；FY 口徑）"&' + lit(b)
        elif v.startswith("估值：目標價＝十年 non-GAAP P/E 中位數"):   # A4
            w = ("估值：目標價＝DCF（標準版：無槓桿自由現金流 FCFF＋WACC）與 Comparable（同業前瞻 P/E 中位數 × 相對倍數調整係數 × 12 個月後前瞻 EPS）加權，DCF 權重 tp_w_dcf（Andy 2026-10-25 定 50％）；"
                 "DCF 與 Comparable 各自以情境權重 A 0.3／B 0.5／C 0.2 加權。舊法（十年 non-GAAP P/E 中位數 × FY2027 EPS、循環位置規則值）保留在 Val 頁，標「舊法，不進結論」（v0.4.5 起；驗收 A4、原則 42）。")
        elif v.startswith("毛利率：GPU 毛利率") and "依 GW 線性歸屬" in v and "w" in P46:
            w = v.replace("依 GW 線性歸屬，Assumed", "閘門式歸屬：出貨 GW、到期日〔OpenAI 2030-10-05、Meta 2031-02-23〕、股價門檻與技術商業條件；v0.4.6 起，原則 37")
        elif v.startswith("已知限制：") and "(5) 情境 B、C 只影響加速器顆數與 HBM 價格，CPU 等業務三情境相同。" in v and j2:
            w = v.replace("(5) 情境 B、C 只影響加速器顆數與 HBM 價格，CPU 等業務三情境相同。", "(5) 情境 A／C 自 FY2027 起帶 AMD 顆數係數與資料中心 CPU 成長路徑（v0.4.6，J2）；Client、Gaming、Embedded、稅率三情境相同（無機制或證據，X27）。")
            w = w.replace("本模型只計股數稀釋。", "本模型只計股數稀釋（營收減項已含在報表營收，見 DCF 頁調節橋）。")
        elif v.startswith("情境對應（v0.2 延伸期）") and j2:
            w = v.replace("FY2027 三情境相同，情境差從 FY2028 起。", "v0.4.6 起情境 A／C 自 FY2027 起另帶 AMD 顆數係數（×1.2／×0.8）與資料中心 CPU 成長（高／低）路徑（驗收 J2）。")
        elif v.startswith("重點觀察 2：Q3 實績營收"):   # A6：手寫舊數字改公式
            a = v.split("FY2027 EPS 9.14 → 11.02")[0]
            w = ("=" + lit(a) + f'&"FY2027 EPS（情境 B）"&TEXT({eps("B", "E")},"0.00")&"，比共識 15.76（stockanalysis，2026-10-23，財報前；FY 口徑）低 "&TEXT(1-{eps("B", "E")}/cons_eps27,"0.0%")&"；'
                 f'12 個月後前瞻 EPS（情境加權）"&TEXT({feps},"0.00")&" 比同時間窗共識 "&TEXT(2/12*cons_eps27+10/12*cons_eps28,"0.00")&" 低 "&TEXT(1-{feps}/(2/12*cons_eps27+10/12*cons_eps28),"0.0%")&"（分母：2/12 × 共識 FY2027＋10/12 × 共識 FY2028 約 20，二手）。"')
        elif v.startswith("重點觀察 3：Q4 資料中心＝指引 − 其他分部"):
            b = v.split("下一個檢驗點：", 1)[1]
            w = ('="重點觀察 3：Q4 資料中心＝指引 − 其他分部，推得 "&TEXT(Q!J5,"0.00")&"（季增 "&TEXT(Q!J5/Q!I5-1,"0%")&"）；GPU／CPU 拆分仍為 Assumed，Q4 GPU 推得約 "&TEXT(Q!J6,"0.0")&"（季增約 "&TEXT(Q!J6/Q!I6-1,"0%")&"），k_rec 對這個拆分最敏感（Sens 業務主因表有 GPU 占比 40%／60% 與 Q4 其他分部 ±0.5／±1.0 案例）。下一個檢驗點："&' + lit(b))
        elif v.startswith("="):
            if '命題 4：股價約為共識 FY2027 EPS 的 "' in v:
                w = v.replace('&" 倍；', '&" 倍（分母：stockanalysis 共識 FY2027 EPS 15.76，2026-10-23，財報前；FY 口徑）；')
            if "重點 1：加權目標價約" in v and j2:
                w = w.replace("（三情境相同，時間窗約等於未來 12 個月）", "（情境 B，時間窗約等於未來 12 個月）")
            if "重點 2：FY2027 三情境相同" in v and j2:
                w = w.replace('="重點 2：FY2027 三情境相同（EPS "&TEXT(' + eps("B", "E") + ',"0.00")&" 美元），分歧從 FY2028 開始：',
                              '="重點 2：FY2027 EPS 情境 A（樂觀）／B（基準）／C（保守）為 "&TEXT(' + eps("A", "E") + ',"0.00")&"／"&TEXT(' + eps("B", "E") + ',"0.00")&"／"&TEXT(' + eps("C", "E") + ',"0.00")&" 美元（v0.4.6 起情境 A／C 自 FY2027 起帶 AMD 顆數與 CPU 路徑）；')
            if j2:
                w = w.replace("以情境 A 為基準", "以情境 B 為基準").replace("敏感度以情境 A 為基準（FY2028 起吸收比例 \"&TEXT(abs_mem_A", "敏感度以情境 B 為基準（FY2028 起吸收比例 \"&TEXT(abs_mem_B")
            if "資料中心營收能否「2027 年逾一倍」" in v:   # A9：條件式
                if j2:
                    w = w.replace("Recon!C7", "Recon!E7").replace("Recon!D7", "Recon!F7").replace("Recon!C14", "Recon!E14").replace("Recon!C22", "Recon!E22").replace("（十億美元，三情境相同）", "（十億美元，情境 B）")
                w = w.replace('答：本模型差一點。', '答：在情境 B 的路徑下（AMD 顆數、世代售價係數 1.0、HBM 吸收比例為基準假設），本模型差一點；這是一條路徑的結果，不是機率結論。')
            if "EPS 撐得住現價嗎？" in v and j2:
                w = w.replace("Recon!C10", "Recon!E10").replace("Recon!D10", "Recon!F10").replace('FY2027 EPS "&TEXT(Recon!E10', 'FY2027 EPS（情境 B）"&TEXT(Recon!E10')
            if "EPS 撐得住現價嗎？" in v:
                w = w.replace('&"（另一共識來源為 "', '&"（分母：stockanalysis 共識，2026-10-23，財報前；另一共識來源為 tickflow 四季合計 "')
                w = w.replace('現價約為情境 B FY2028 EPS 的 "', '現價約為本模型情境 B FY2028 EPS（FY 口徑）的 "')
            if "認股權證會稀釋多少？" in v and "w" in P46:
                w = w.replace('"認股權證會稀釋多少？答：依已歸屬 GW 換算，', '"認股權證會稀釋多少？答：依閘門式歸屬（出貨 GW、到期日 OpenAI 2030-10-05／Meta 2031-02-23、股價門檻與條件 wf "&TEXT(wf,"0%")&"）換算，')
                w = w.replace('&"。"', '&"。評價用認股權證股數 "&TEXT(w_val,"0.000")&" 十億股（上限 "&TEXT(w_tot,"0.00")&"；到期前約可歸屬 "&TEXT(w_val/w_tot,"0%")&"）。"', 1) if w.endswith('&"。"') else w
            if "重點 4：v0.4 起 DCF 改為標準版" in v:
                if "w" in P46:
                    w = w.replace('評價股數含客戶認股權證全額 "&TEXT(w_tot,"0.00")&" 十億股。"', '評價股數含客戶認股權證 "&TEXT(w_val,"0.000")&" 十億股（上限 "&TEXT(w_tot,"0.00")&"；到期日前歸屬 GW × wf，原則 37）。"')
                if "sbc" in P46:
                    w = w.replace('（營收的 "&TEXT(sbc_rev,"0.0%")&"）', '（營收的 "&TEXT(sbc_rev,"0.0%")&"，最近四季）')
        if w != v: RD.cell(r, 1).value = w
    rr = lab(RD, '="投資建議：')
    RD.cell(rr, 1).value = (f'="投資建議："&{rec}&"（依上檔空間的估值評等；未評估借券、流通量與時點；加權目標價上檔空間 "&TEXT({up},"+0.0%;-0.0%")&"）。'
                            f'承重假設：DCF 權重 "&TEXT(tp_w_dcf,"0%")&"（權重 0.3 時評等為 "&IF(Sens!C{lab(wb["Sens"], "DCF 權重 0.3")}/px-1>=-0.25,"中立","放空")&"）；GPU 認列售價延續 FY2026 校準（世代售價係數 1.0）；HBM 漲價吸收比例、GPU 毛利率 50％、FY2026 GPU 拆分為假設；'
                            f'Comparable 相對倍數係數 "&TEXT(rm_factor,"0.00")&"（已達夾值上限）。主要失效條件見本頁末。"')
    # 版本沿革列
    h = lab(RD, "v0.4.4：")
    RD.insert_rows(h)
    RD.cell(h, 1, f"{VER}（{BUILT}）：AMD v0.4.4 驗收報告採用項（Andy 2026-11-10 07:52 確認採用規劃書，J1–J4 照建議預設）。" +
            ("數值：認股權證改閘門式（出貨 GW 隨 AMD 顆數、到期日失效、wf 股價與條件；PL 與 DCF 評價股數同一邏輯，J1a）；情境 A／C 自 FY2027 起帶 AMD 顆數與資料中心 CPU 路徑（J2b），EPS 敏感度改以情境 B 為基準；"
             "目標時點改現價日＋12 個月；評價股數改流通股＋庫藏股法；股票薪酬比率改最近四季；目標價 Sens 新增業務主因整本重算表（GPU 占比 40／60％、Q4 其他分部 ±0.5／±1.0、認股權證 0／50％、不夾值、現金稅率）。目標價變動拆解見下一列。"
             if ctx["V46"] else
             "展示與治理（數值不變）：build_id、輸入指紋與模型狀態；In、Q 頁單一主要標記與來源誘因；歷史財務移到 Hist 頁；三個日期分列；舊法移出主敘述；倍數標分母；手寫舊數字改公式；命題 1 條件式；"
             "DCF 頁 non-GAAP → FCFF 調節橋與現金稅率；Recon 反解；資料基準日與逾期；風險登錄、失效條件、可證偽指標、評等邊界；目標價 Sens 加 DCF 單法美元欄；β 依據。"))
    if ctx["V46"] and ctx.get("ATTR"):
        RD.insert_rows(h + 1); RD.cell(h + 1, 1, ctx["ATTR"])
    # 新增各段（本頁末）
    r = RD.max_row + 2
    def put(t, bold=False):
        nonlocal r
        RD.cell(r, 1, t).font = Font(bold=bold); r += 1
        return r - 1
    G["h"] = put(f"{VER} 驗收採用項（HTML 用；AMD v0.4.4 驗收報告，採用規劃書 2026-11-10）", True)
    G["state"] = put(f'="build_id "&Sens!D{G["gov"]["bid"]}&"；模型狀態 "&Sens!D{G["gov"]["st"]}&"（Sens 業務主因表與 HTML 為建置快照：只在 Excel 改 In 頁而未重跑 python3 amd/sens_amd.py 與建置時顯示 STALE）"')
    u = up; e = f"IF({u}<-0.375,-0.5,IF({u}<0,-0.25,IF({u}<0.375,0.25,0.5)))"; d = f"ABS({u}-{e})"
    G["edge"] = put(f'=IF({d}<0.05,"評等貼近級距邊界（原則 36）：加權目標價上檔空間 "&TEXT({u},"+0.0%;-0.0%")&" 距 "&TEXT({e},"+0%;-0%")&" 邊界只有 "&TEXT({d}*100,"0.0")&" 個百分點；加權目標價 "&TEXT(px*(1+{e}),"0.00")&" 美元（目前 "&TEXT({tp},"0.00")&"）即翻轉評等。",'
                    f'"評等邊界（原則 36）：上檔空間距最近級距邊界 "&TEXT({d}*100,"0.0")&" 個百分點（5 個百分點以上，評等不貼近邊界）；加權目標價 "&TEXT(px*(1+{e}),"0.00")&" 美元時翻轉。")')
    G["asof_h"] = put(f"資料基準日（as_of）：{AS_OF}（現價日）。市場資料與基準日相差超過 2 個交易日標「逾期」；最近一次財報（3Q26，2026-11-03）前的共識標「財報前，只作歷史對照」（原則 15）。", True)
    RD.cell(r, 1, "項目").font = BOLD; RD.cell(r, 2, "資料日期").font = BOLD; RD.cell(r, 3, "狀態").font = BOLD; G["asof_t"] = r; r += 1
    items = [("股價 px（Yahoo Finance 日線收盤）", "2026-11-05", 0), ("同業股價（五家；Comparable 前瞻 P/E 換算）", "2026-11-05", 0),
             ("同業前瞻 EPS（stockanalysis 統計頁）", "2026-10-30", 0), ("無風險利率 rf（美國 10 年期公債殖利率）", "2026-10-23", 0),
             ("相對倍數歷史資料（Bloomberg 前瞻倍數，10 年）", "2026-10-29", 0),
             ("市場共識目標價 cons_tp", "2026-10-25", 1), ("共識 FY2026／FY2027 EPS（stockanalysis、tickflow）", "2026-10-23", 1), ("共識 FY2028 EPS（約 20，二手）", "2026-07-30", 1),
             ("資產負債（3Q26；8-K 與 10-Q）", "2026-09-26", 2)]
    G["asof_rows"] = []
    for labl, dd, kind in items:
        RD.cell(r, 1, labl); RD.cell(r, 2, dd)
        RD.cell(r, 3, '="財報前，只作歷史對照"' if kind == 1 else ('="最近一季資產負債表（以 stub 期滾到評價日）"' if kind == 2 else
                       f'=IF(ABS(NETWORKDAYS(MIN(DATEVALUE(B{r}),DATEVALUE("{AS_OF}")),MAX(DATEVALUE(B{r}),DATEVALUE("{AS_OF}")))-1)>2,"逾期 "&TEXT(ABS(NETWORKDAYS(MIN(DATEVALUE(B{r}),DATEVALUE("{AS_OF}")),MAX(DATEVALUE(B{r}),DATEVALUE("{AS_OF}")))-1),"0")&" 個交易日","同期")'))
        G["asof_rows"].append(r); r += 1
    G["asof_n"] = put("說明：rf、同業前瞻 EPS 與相對倍數資料逾期，但與其他 AI 半導體個股模型同一取值日，為維持跨模型一致暫不單獨更新（原則 15；不需驗收清單 X24）；同業 P/E 已以 2026-11-05 收盤換算。共識只作對帳，不進目標價。")
    inv = G["recon"]["inv"]
    G["fail_h"] = put("主要失效條件（可觀察；原則 3）", True)
    E_up, E_dn = f"Recon!B{G['recon']['e_up']}", f"Recon!B{G['recon']['e_dn']}"
    fails = [f'="(1) DCF：DCF 目標價（目前 "&TEXT(Recon!C{inv["dcf_up"]},"0.00")&"）升到約 "&TEXT(Recon!B{inv["dcf_up"]},"0.00")&"，上檔空間到 "&TEXT({E_up},"+0%;-0%")&"、評等升一級（Comparable 不變）；支持現價需約 "&TEXT(Recon!B{inv["dcf_px"]},"0.00")&"。可觀察：WACC 組成（rf、自算產業 β）與 FCFF 路徑（FY2027–29 營收與營業利益率）。"',
             f'="(2) 倍數：Comparable 採用倍數（同業中位數 × 相對倍數係數；目前 "&TEXT({pmed}*rm_factor,"0.0")&" 倍）升到約 "&TEXT(Recon!B{inv["pe_up"]},"0.0")&" 倍評等升一級；支持現價需約 "&TEXT(Recon!B{inv["pe_px"]},"0.0")&" 倍。可觀察：同業前瞻本益比中位數。"',
             f'="(3) 盈餘：12 個月後前瞻 EPS（情境加權；模型 "&TEXT(Recon!C{inv["eps_up"]},"0.00")&" 美元）在同業中位數與係數下需到約 "&TEXT(Recon!B{inv["eps_up"]},"0.00")&" 美元評等才升一級；支持現價需約 "&TEXT(Recon!B{inv["eps_px"]},"0.00")&" 美元。可觀察：4Q26 資料中心營收（2027-02 初）、2027 年「資料中心逾一倍」說法是否維持、FY2027 non-GAAP EPS 共識是否收斂。"']
    if "pk_up" in inv:
        _x = lambda k: f'IF(ISNUMBER(Recon!B{inv[k]}),TEXT(Recon!B{inv[k]},"0.00")&" 倍","超出測試範圍")'
        fails.append(f'="(4) GPU 認列售價：FY2027 起世代售價係數三情境同乘約 "&{_x("pk_up")}&"時評等升一級；支持現價需約 "&{_x("pk_px")}&"。可觀察：MI450／Helios 系統 ASP 與毛利、每季資料中心營收 ÷ 出貨 GW。"')
    fails.append('="(5) 認股權證：各批歸屬的 10-Q 揭露（截至 2026-09-26 尚無一批歸屬）；若股價門檻與條件只達成一半（wf 0.5），評價股數減少、目標價上升，見 Sens 業務主因表。"')
    G["fail"] = [put(t) for t in fails]
    G["frame"] = put(f'="估值框架並列：100% DCF＝"&TEXT({dcf},"0.00")&" 美元、100% Comparable（相對倍數調整後）＝"&TEXT({cmp_},"0.00")&" 美元、加權（DCF "&TEXT(tp_w_dcf,"0%")&"）＝"&TEXT({tp},"0.00")&" 美元；DCF 情境價格帶 "&TEXT(Val!B{band},"0.0")&"～"&TEXT(Val!C{band},"0.0")&" 美元；只看 Comparable（未調整）"&TEXT(Val!B{lab(V, "Comparable 目標價（市場主流同業法）")},"0.00")&" 美元。"')
    G["risk_h"] = put("風險登錄（原則 10、19；是否建模、理由、觀察指標）", True)
    G["risk"] = [put(t) for t in (
        "客戶認股權證（OpenAI、Meta 各 160M 股）：建模（閘門式；出貨 GW、到期日、wf）。理由：合約 8-K 已揭露上限、首批條件、最後一批 $600 與到期日；中間門檻與各批股數未揭露，只以 wf 0／50％ 敏感度呈現（X19）。觀察：10-Q 認股權證附註的已歸屬批數與公允價值。",
        "大單時點（OpenAI、Meta、Anthropic、Oracle）：建模於 AMD 顆數係數（情境 A／C 自 FY2027 起 ×1.2／×0.8，及三情境同乘的目標價 Sens）。觀察：首 1 GW MI450 出貨與認列時點、客戶部署里程碑公告。",
        "ROCm 軟體採用與系統級議價（Helios 機櫃）：未建模（無公開的採用率與議價資料，拆成驅動屬 X19）。理由：效果已含在 k_rec 與世代售價係數的區間。觀察：主要雲端與模型商公開採用 ROCm 的工作負載、Helios 系統 ASP。",
        "CoWoS／HBM 供給上限：部分建模（加速器顆數讀 Siliconomics／markets/memory 共用路徑，HBM 價格路徑三情境）；AMD 自身的封裝與 HBM 配額未建模（未揭露）。觀察：台積電 CoWoS 產能分配、HBM4 供應商合約。",
        "HBM 漲價轉嫁：建模於吸收比例（情境 A／B／C 0.2／0.5／0.7 與 0／1 Sens）。觀察：MI450 報價與毛利率指引。",
        "GPU／CPU 拆分（公司不揭露）：建模於 FY2026 GPU 錨點與 Sens 40％／60％（整本重算 k_rec）。觀察：公司任何 Instinct 營收揭露、10-K 分部說明。",
        "流通量、借券與軋空：不進目標價（估值評等不含交易執行；不需驗收清單 X14）。觀察：借券費率、空單比例。")]
    G["ind_h"] = put("主要假設驅動與可證偽的公開指標（原則 20）", True)
    G["ind"] = [put(t) for t in (
        "FY2026 GPU 拆分（gs_q1、gs_q2、ng_q3、ng_q4，Assumed）：公司 Instinct 營收揭露；4Q26 資料中心營收（2027-02 初）。",
        "世代售價係數（pg_*，Analogy）：MI450／Helios 系統 ASP、每季資料中心營收 ÷ 出貨 GW。",
        "HBM 吸收比例（abs_mem_*，Assumed）：資料中心毛利率與公司毛利率指引。",
        "GPU 毛利率（gm_gpu，Assumed）：公司「資料中心 AI 毛利率低於平均」說法的量化揭露。",
        "AMD 顆數係數（uf、uf_A／uf_C，Assumed）：首 1 GW 出貨時點、各客戶 GW 里程碑公告。",
        "認股權證歸屬（gw_*、wf，Assumed）：10-Q 認股權證附註的已歸屬批數；AMD 股價相對各批門檻（最後一批 $600）。",
        "資料中心 CPU 成長（g_cpu*，Assumed）：公司伺服器 CPU 年增說法（2027 逾 70％）在季度營收的實現。")]
    return r


def apply(wb, here, ctx):
    V46, OVR = ctx["V46"], ctx["OVR"]
    sj_path = os.path.join(here, "sens.json")
    SJ = json.load(open(sj_path, encoding="utf-8")) if (os.path.exists(sj_path) and not OVR) else None
    if SJ and SJ.get("_stage") != ctx["STAGE"]: SJ = None
    hist_sheet(wb); hist_links(wb, ctx["V46P"], OVR)
    if V46: v46_values(wb, ctx["V46P"], OVR)
    q_overrides(wb, OVR)
    n = norm_tags(wb)
    j2 = "j2" in ctx["V46P"]
    PLs = wb["PL"]
    PLs["A1"].value = ("PL — AMD 損益（US$B；non-GAAP；FY2025A 實績、FY2026 錨點，FY2027 起為模型）。會計年度＝日曆年。三情境各一個完整推導區塊（本頁下方），頂端為摘要。"
                       + ("情境 A／C 自 FY2027 起帶 AMD 顆數係數與資料中心 CPU 成長路徑（驗收 J2），產業顆數係數差異自 CY2028 起" if j2 else "FY2026、FY2027 三情境相同（情境係數從 CY2028 才開始）"))
    # A3：rf 期限與日期；B6：β 依據；B7：承重外部輸入定位
    in_set(wb, "rf", 9, str(wb["In"].cell(inrow(wb, "rf"), 9).value or "") + "。期限：10 年期；取值日 2026-10-23（股價日 2026-11-05，逾期 9 個交易日，與其他個股模型同一取值日；見 README 資料基準日表）")
    for k in ("beta", "beta_u"):
        c = wb["In"].cell(inrow(wb, k), 9)
        c.value = str(c.value or "") + "。依據：docs/plan/自算產業β_20261105.md（US 半導體 13 家；AMD 本身在同業組內）"
    for k, add in (("w_tot", "定位：OpenAI 8-K Item 1.01 與 Exhibit 4.1（up to 160 million shares、$0.01、首批隨首 1 GW、最後一批 $600）；Meta 8-K 2026-02-23 Item 1.01（同結構）。轉換：0.16＋0.16＝0.32 十億股"),
                   ("q4_rev", "定位：3Q26 新聞稿「Current Outlook」段：approximately $17.0 billion, plus or minus $500 million（區間以 ±3% 表示）"),
                   ("px", "定位：Yahoo Finance AMD 日線 2026-11-05 Close 404.40（未調整）")):
        c = wb["In"].cell(inrow(wb, k), 9); c.value = str(c.value or "") + "。" + add
    bid = build_id(here, ctx["STAGE"])
    G = {}
    G["gov"] = governance(wb, SJ, bid, OVR)
    G["sens"] = sens_tp(wb, SJ, V46)
    G["recon"] = recon(wb, SJ, V46)
    G["val"] = val(wb, ctx)
    G["bridge"] = dcf_bridge(wb)
    if V46 and SJ and SJ.get("attr"): ctx["ATTR"] = SJ["attr"]
    readme(wb, ctx, bid, G)
    SE = wb["Sens"]
    SE["A1"].value = (f"Sens — 一次動一個輸入（情境 {ctx['SB0']}）：對 FY2027、FY2028、FY2029 的營收與 EPS 的影響；「舊法隱含價」欄＝EPS × 十年 non-GAAP P/E 中位數（舊法，只作參考，不是目標價的方法）。"
                      "每個案例是一個完整的推導區塊（本頁下方）。目標價的敏感度見下方「目標價參數敏感度」與「業務主因（整本重算）」兩表；模型狀態與 build_id 見本頁底部")
    for c in (6, 7, 11, 12):
        v = SE.cell(3, c).value
        if isinstance(v, str) and "隱含價" in v: SE.cell(3, c).value = v.replace("隱含價（中位數 P/E）", "舊法隱含價（參考）").replace("隱含價相對基準", "舊法隱含價相對基準")
    rn = lab(SE, "讀法：")
    SE.cell(rn, 1).value = (("讀法：v0.4.6 起以情境 B 為基準（情境 A／C 自 FY2027 起帶 AMD 顆數與 CPU 路徑）；顆數、售價、吸收比例與 CPU 成長從 FY2027 起作用。" if j2 else
                             "讀法：FY2027 三情境相同；顆數、售價、吸收比例與 CPU 成長從 FY2027 起作用，情境 B、C 從 FY2028 起。")
                            + "舊法隱含價欄＝EPS × 十年 non-GAAP P/E 中位數（舊法，不進結論）。世代售價係數案例的 GPU 成本與營收等比例（毛利率不變）：售價 ×1.5 時每顆成本增加約為原售價的 25％，大於 HBM 容量增加 144 GB 的成本（約 9％），屬保守處理，不另做「只動售價」案例（驗收 A7 核對）。"
                            + ("認股權證歸屬隨 AMD 顆數係數變動（v0.4.6 起，原則 37）。" if "w" in ctx["V46P"] else ""))
    return f"v045_post（{ctx['STAGE']}）：標記正規化 {n} 格；build_id {bid}；快照 {'有' if SJ else '無'}"
