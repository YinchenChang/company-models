# 舊範本治理項目共用補齊函式（外部稽核優化 W3-6，2026-11-23；J4：一次補齊、不改值、不升主版）。
#
# 建置腳本在 wb.save() 之前呼叫一次：
#     sys.path.insert(0, <repo>/tools/model_governance); import stamp_governance as SG
#     print(SG.apply(wb, here=<公司資料夾>, slug="cohr", ver="0.4.7", snapshot_tp=<sens.json 的基準加權目標價或 None>,
#                    sens_keys={<Sens 頁 A 欄案例文字>: "鍵1 鍵2"}, exempt={<鍵>: <不列的理由>}))
#
# 補的是 tools/model_governance/check_governance.py 與 tools/sens_coverage/check_sens_coverage.py 檢查的項目，做法照已通過的
# spacex/build_spacex.py（BUILD_ID、norm_tag、Sens 指紋三段）與 amd/v045_post.py（norm_tags、build_id、governance），只是抽成共用：
#   1. build_id：＝<slug>-v<版號>-<公司資料夾內建置腳本與資料檔的雜湊前 10 碼>；寫在 Sens 頁底「build_id」列 D 欄、README!A1 末尾。
#      HTML 由 tools/company_html/build_company_html.py 頁尾自動帶出（Sens 頁有 build_id 列而頁面內文沒有時）。
#   2. 模型狀態：Sens 頁底「輸入指紋」兩列與「模型狀態」列。指紋只算 In 頁低／基準／高欄的「手填數字」（不含公式格）：
#      建置時由本函式算出寫死一份，活頁簿內另有一格即時重算；有人只在 Excel 改 In 頁藍字而沒重跑建置，兩者不同 → STALE。
#      有整本重算快照（sens.json）的模型再加一個條件：加權目標價與快照基準相同（snapshot_tp）。
#   3. 單一主要標記與「來源誘因」欄（In、Q 頁）：混標（例「Verified／Interested-party（指引）」）改成一個主要標記，原混標文字寫進來源誘因欄；
#      Interested-party 列補誘因文字。主要標記的取法見 primary_tag()（原則第 10 節第 2 條：以最近一期原文的標記為主）。
#   4. Sens 鍵欄：(a) tools/fcff_dcf.py 的「目標價參數敏感度」表自己有「鍵」欄，但覆蓋檢查只讀頁面第一個表頭的鍵欄；兩者不同欄時，
#      把鍵附在該列 B 欄（設定）文字後面（「；鍵：wacc_adj」），檢查的第 3 條認 A、B 欄的獨立識別字。
#      (b) sens_keys：整本重算快照表等沒有鍵的案例列，依呼叫端給的對照寫進鍵欄（鍵欄已有內容或被數字占用時附在 B 欄）。
#      (c) exempt：確實不進目標價的鍵，在 Sens 頁底寫「<鍵> 不列：<理由>」。理由由各家模型負責人寫，本函式不代寫。
# 本函式只新增列、欄與文字，不動任何既有數值格與公式（B 欄設定文字、In／Q 頁標記文字除外）。
import datetime, glob, hashlib, os, re
from openpyxl.styles import Font, PatternFill

TAGS = ("Verified", "Interested-party", "Analogy", "Assumed", "Derived", "—")
BOLD = Font(bold=True); HDR = PatternFill("solid", fgColor="DDEBF7")
HASH_GLOBS = ("build_*.py", "*_post.py", "company.py", "sens.json", "segments.json", "hist_frozen*.json", "coord_alt.json")


def build_id(here, slug, ver, extra_files=()):
    """版本＋建置輸入雜湊。雜湊範圍＝公司資料夾內的建置腳本、後處理補丁與資料 json（不含 html.json 等由成品產生的檔）。"""
    h = hashlib.sha256(ver.encode())
    files = sorted({p for g in HASH_GLOBS for p in glob.glob(os.path.join(here, g))} | {os.path.join(here, f) for f in extra_files})
    for p in files:
        if os.path.isfile(p):
            h.update(os.path.basename(p).encode()); h.update(open(p, "rb").read())
    return f"{slug}-v{ver.lstrip('v')}-{h.hexdigest()[:10]}"


def incentive(text):
    """Interested-party 列的來源誘因（照 amd/v045_post.py 的分類）。"""
    t = str(text or "")
    if any(k in t for k in ("共識", "賣方", "分析師", "Yahoo", "stockanalysis", "StockAnalysis", "MarketBeat", "Benzinga", "TIKR", "目標價")):
        return "賣方分析師或資料商：評等與交易業務誘因，目標價與預估傾向跟隨股價"
    if any(k in t for k in ("法說", "逐字稿", "二手", "媒體")):
        return "公司管理層口頭陳述（含二手轉述）：股價與市場預期管理誘因，未經查核、可能偏樂觀"
    if any(k in t for k in ("指引", "公司", "新聞稿", "8-K", "簡報", "年報")):
        return "公司指引或自揭數字：預期管理誘因，口徑由公司定義"
    return "利害關係方提供：未經獨立查核"


def primary_tag(t):
    """混標 → 單一主要標記。規則（原則第 10 節第 2 條「以最近一期原文的標記為主」）：
    混標各段依欄位由舊到新排列，取最後一段的標記；任何一段是 Interested-party（指引、公司說法）時取 Interested-party（與 amd、alab 的既有做法相同）。"""
    parts = [p.strip() for p in re.split(r"／|/|、", re.sub(r"（[^）]*）", "", t)) if p.strip()]
    found = [next((g for g in TAGS[:5] if p.startswith(g)), None) for p in parts]
    found = [f for f in found if f]
    if not found:
        return None
    return "Interested-party" if "Interested-party" in found else found[-1]


def norm_tags(wb, override=None):
    """In、Q 頁：加「來源誘因」欄、混標改單一主要標記。override＝{(工作表, 列號或 A 欄鍵): 標記} 可逐列指定。回傳改動格數。"""
    override = override or {}; n = 0; skipped = []
    for sh in ("In", "Q"):
        if sh not in wb.sheetnames: continue
        ws = wb[sh]; hr = tcol = None
        for r in range(1, 15):
            vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
            if "標記" in vals: hr, tcol = r, vals.index("標記") + 1; break
        if hr is None: skipped.append(f"{sh} 頁找不到「標記」欄"); continue
        vals = [ws.cell(hr, c).value for c in range(1, ws.max_column + 1)]
        scol = vals.index("來源") + 1 if "來源" in vals else None
        icol = next((i + 1 for i, v in enumerate(vals) if isinstance(v, str) and "來源誘因" in v), None)
        if icol is None:
            icol = max(i + 1 for i, v in enumerate(vals) if v is not None) + 1
            c = ws.cell(hr, icol, "來源誘因"); c.font = BOLD; c.fill = HDR; ws.column_dimensions[c.column_letter].width = 50
        for r in range(hr + 1, ws.max_row + 1):
            t = ws.cell(r, tcol).value
            if not isinstance(t, str) or t == "標記": continue
            t0 = t
            if t not in TAGS:
                prim = override.get((sh, r)) or override.get((sh, ws.cell(r, 1).value)) or primary_tag(t)
                if prim is None: skipped.append(f"{sh}!{ws.cell(r, tcol).coordinate}「{t}」"); continue
                ws.cell(r, tcol).value = prim; n += 1
                old = ws.cell(r, icol).value
                ws.cell(r, icol).value = (f"{old}；" if old else "") + f"原混標「{t}」：各欄來源不同，主要標記取最近一期（或公司指引）的標記"
                t = prim
            if t == "Interested-party":
                cur = str(ws.cell(r, icol).value or "")
                if not cur or cur.startswith("原混標"):
                    src = f"{ws.cell(r, scol).value if scol else ''}｜{ws.cell(r, 1).value or ''}｜{ws.cell(r, 2).value if isinstance(ws.cell(r, 2).value, str) else ''}｜{t0}"
                    ws.cell(r, icol).value = incentive(src) + (f"；{cur}" if cur else ""); n += 1
    return n, skipped


def _in_range(wb):
    """指紋範圍：In 頁表頭下一列到「情境驅動對照」前一列（沒有該節就到最後一列）的 D:F。"""
    IN = wb["In"]
    hr = next(r for r in range(1, 15) if "標記" in [IN.cell(r, c).value for c in range(1, IN.max_column + 1)])
    last = next((r for r in range(hr + 1, IN.max_row + 1) if str(IN.cell(r, 1).value or "").startswith("情境驅動對照")), IN.max_row + 1) - 1
    return hr + 1, last


def _serial(v):
    if isinstance(v, datetime.datetime): return (v - datetime.datetime(1899, 12, 30)).total_seconds() / 86400
    if isinstance(v, datetime.date): return (v - datetime.date(1899, 12, 30)).days
    return v


def fingerprint(wb):
    """In 頁 D:F 手填數字（不含公式、文字）× (列 × 3＋欄) 的加總；與活頁簿內的即時公式同一算法。"""
    IN = wb["In"]; r0, r1 = _in_range(wb); s = 0.0
    for r in range(r0, r1 + 1):
        for c in (4, 5, 6):
            v = _serial(IN.cell(r, c).value)
            if isinstance(v, (int, float)) and not isinstance(v, bool): s += v * (r * 3 + c)
    return s


def _lab(ws, prefix):
    return next((r for r in range(1, ws.max_row + 1) if isinstance(ws.cell(r, 1).value, str) and ws.cell(r, 1).value.strip().startswith(prefix)), None)


def stamp(wb, bid, snapshot_tp=None, rerun_hint="重跑建置腳本"):
    """Sens 頁底加輸入指紋、模型狀態、build_id 三段；README!A1 末尾加 build_id。回傳各列列號。"""
    SE = wb["Sens"]; r0, r1 = _in_range(wb); rng = f"In!$D${r0}:$F${r1}"
    for pre in ("build_id", "模型狀態"):
        assert _lab(SE, pre) is None, f"Sens 頁已有「{pre}」列，不重複加"
    r = SE.max_row + 3; G = {}
    SE.cell(r, 1, "模型治理（外部稽核優化 W3-6，2026-11-23）：HTML、README 與 Sens 頁的整本重算快照是建置時的結果。只在 Excel 改 In 頁藍字而沒有" + rerun_hint + "時，下方模型狀態顯示 STALE").font = BOLD; r += 1
    SE.cell(r, 1, "輸入指紋（本檔即時重算；In 頁低／基準／高欄的手填數字，不含公式格）")
    SE.cell(r, 4, f"=SUMPRODUCT(--NOT(_xlfn.ISFORMULA({rng})),{rng},ROW({rng})*3+COLUMN({rng}))").number_format = "0.000000"; G["fp"] = r; r += 1
    SE.cell(r, 1, "輸入指紋（建置時）"); SE.cell(r, 4, fingerprint(wb)).number_format = "0.000000"; G["fp0"] = r; r += 1
    cond = [f"ABS(D{G['fp']}-D{G['fp0']})<=0.000000001*MAX(1,ABS(D{G['fp0']}))"]
    if snapshot_tp is not None:
        V = wb["Val"]; tr = _lab(V, "目標價（DCF 與 Comparable 加權）")
        assert tr, "Val 頁找不到「目標價（DCF 與 Comparable 加權）」列"
        SE.cell(r, 1, "加權目標價（整本重算快照的基準檔）"); SE.cell(r, 4, snapshot_tp).number_format = "#,##0.00"; G["tp0"] = r; r += 1
        cond.append(f"ABS(Val!$B${tr}-D{G['tp0']})<0.005")
    SE.cell(r, 1, f"模型狀態（CURRENT＝Sens 快照、HTML、README 與本檔輸入一致；STALE＝In 頁已改、須{rerun_hint}）").font = BOLD
    SE.cell(r, 4, f'=IF(AND({",".join(cond)}),"CURRENT","STALE")'); G["st"] = r; r += 1
    SE.cell(r, 1, "build_id（README、Sens、HTML 三處相同）"); SE.cell(r, 4, bid); G["bid"] = r
    a1 = wb["README"]["A1"]
    if isinstance(a1.value, str) and a1.value.startswith("="): a1.value = a1.value + f'&"；build_id {bid}"'
    else: a1.value = f"{a1.value}；build_id {bid}"
    return G


def _first_header(SE):
    """與 check_sens_coverage.sens_region 相同的表頭判斷：前 40 列第一個至少三格非空、且有 20 字以內「…目標價」欄名的列。"""
    for r in range(1, min(SE.max_row, 40) + 1):
        vals = [str(SE.cell(r, c).value or "").strip() for c in range(1, min(SE.max_column, 30) + 1)]
        if sum(bool(v) for v in vals) >= 3 and any("目標價" in v and len(v) <= 20 for v in vals):
            return r, next((i + 1 for i, v in enumerate(vals) if v.startswith("鍵")), None)
    return None, None


def _append_text(cell, text):
    """在文字格後面接一段字；格內是公式（回傳文字的公式）時用 & 接上。數字格不動，回傳 False。"""
    v = cell.value
    if v is None: cell.value = text.lstrip("；"); return True
    if not isinstance(v, str): return False
    cell.value = (v + f'&"{text}"') if v.startswith("=") else (v + text)
    return True


def _put_key(SE, r, kc, keys):
    """把鍵寫進第 r 列：鍵欄空白就寫鍵欄，否則附在 B 欄（設定）文字後；B 欄是數字時不動（回傳 None）。"""
    if kc and SE.cell(r, kc).value in (None, ""):
        SE.cell(r, kc).value = keys; return "鍵欄"
    return "B 欄" if _append_text(SE.cell(r, 2), f"；鍵：{keys}") else None


def sens_keycol(wb, sens_keys=None, exempt=None):
    """補 Sens 鍵（說明見檔頭第 4 點）。回傳（附鍵列數, 找不到的案例標籤, 說明文字）。"""
    SE = wb["Sens"]; hr, kc = _first_header(SE); n = 0; miss = []
    for r in range(1, SE.max_row + 1):                      # (a) fcff_dcf「目標價參數敏感度」表
        if SE.cell(r, 1).value != "參數": continue
        vals = [str(SE.cell(r, c).value or "") for c in range(1, 15)]
        if "鍵" not in vals: continue
        k2 = vals.index("鍵") + 1
        if hr == r or (hr is not None and kc == k2 and r > hr): continue     # 這張表就是第一個表頭，或鍵欄同欄：檢查本來就讀得到
        rr = r + 1
        while SE.cell(rr, 1).value and not str(SE.cell(rr, 1).value).startswith("檢查"):
            k = SE.cell(rr, k2).value
            if isinstance(k, str) and re.fullmatch(r"[A-Za-z0-9_]+", k) and f"鍵：{k}" not in str(SE.cell(rr, 2).value or ""):
                n += _append_text(SE.cell(rr, 2), f"；鍵：{k}")
            rr += 1
    for lab_, keys in (sens_keys or {}).items():            # (b) 呼叫端給的案例 → 鍵
        rows = [r for r in range((hr or 0) + 1, SE.max_row + 1) if isinstance(SE.cell(r, 1).value, str) and SE.cell(r, 1).value.strip() == lab_.strip()]
        if not rows: miss.append(lab_); continue
        ok = sum(bool(_put_key(SE, r, kc, keys)) for r in rows); n += ok     # 同一標籤出現在多張表時，能寫的列都寫（引用表的 B 欄是數字，略過）
        if not ok: miss.append(lab_ + "（鍵欄與 B 欄都不能寫）")
    if exempt:                                              # (c) 豁免句（理由由呼叫端提供）
        r = SE.max_row + 2
        SE.cell(r, 1, "不列入敏感度案例的輸入（外部稽核優化 W3-6；每句＝鍵＋不列＋理由，敏感度覆蓋檢查第 5 條）").font = BOLD
        for k, why in exempt.items():
            r += 1; SE.cell(r, 1, f"{k} 不列：{why}")
    return n, miss, ("Sens 頁前 40 列找不到含「目標價」欄名的表頭（覆蓋檢查會失敗，須把目標價敏感度表移到頁首或另案處理）" if hr is None else "")


def apply(wb, here, slug, ver, snapshot_tp=None, sens_keys=None, exempt=None, tag_override=None, extra_files=(), rerun_hint="重跑建置腳本"):
    """一次補齊（存檔前呼叫）。順序：鍵與豁免 → 標記與來源誘因 → 指紋、狀態、build_id（放最後，build_id 列在 Sens 頁最底）。"""
    nk, miss, note = sens_keycol(wb, sens_keys, exempt)
    nt, skipped = norm_tags(wb, tag_override)
    bid = build_id(here, slug, ver, extra_files)
    stamp(wb, bid, snapshot_tp, rerun_hint)
    return (f"stamp_governance：build_id {bid}；標記與來源誘因 {nt} 格；Sens 附鍵 {nk} 列"
            + (f"；找不到案例列：{'、'.join(miss)}" if miss else "") + (f"；未處理標記：{'、'.join(skipped)}" if skipped else "") + (f"；{note}" if note else ""))
