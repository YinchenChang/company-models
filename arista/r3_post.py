# 第三輪修正（合併前）共用後處理：建置腳本在存檔前呼叫 apply(wb, …)。只改標籤、格式與 README 舊文，不改任何數值公式。
# 各公司資料夾放同一份（公司自有，不屬於 tools/）；項目編號對照第三輪修正清單。
#   15／23：情境名稱統一為「情境 A（順利）／B（平平）／C（不如預期）」；In 頁「情境驅動對照」收斂法列名改中性、% 與 pt 列格式化
#   24：Val「上檔空間級距」列名註明「六級，供分組」，投資建議用五級規則
#   22：README 頁舊「重點觀察」——重點觀察 1 只留公式部分（刪 2026-10-29 調整前說明），重點觀察 2、3 舊文刪除
#   14：stub 期列名、說明依實際口徑（公司腳本傳入）
#   30：工具字樣（抓取）與本機路徑改中性寫法
#   16：g 受 rf 上限時 Sens「永續成長率 g ±1pt」列名改「永續成長率 g −1pt／上限 rf」
SCN_REPL = (("情境 A（樂觀）", "情境 A（順利）"), ("情境 B（基準）", "情境 B（平平）"), ("情境 C（保守）", "情境 C（不如預期）"),
            ("情境 C（悲觀）", "情境 C（不如預期）"),
            ("情境 A 隱含價（樂觀）", "情境 A 隱含價（順利）"), ("情境 B 隱含價（基準）", "情境 B 隱含價（平平）"), ("情境 C 隱含價（下檔）", "情境 C 隱含價（不如預期）"),
            ("情境權重 A（樂觀）", "情境權重 A（順利）"), ("情境權重 B（基準）", "情境權重 B（平平）"), ("情境權重 C（保守）", "情境權重 C（不如預期）"),
            ("情境權重 C（悲觀）", "情境權重 C（不如預期）"),
            # 30：工具字樣與本機路徑
            ("（chart API 月資料；2026-11-05 抓取）", "（chart API 月資料；2026-11-05 讀取）"), ("2026-11-05 抓取", "2026-11-05 讀取"), ("已逐份抓取", "已逐份核對"),
            ("/home/claude/beta/industry_beta.md", "自算產業 β（2026-11-05，見 docs）"))
OBS_NOTE = "重點觀察 2、3：舊版文字（含過時數字）已於第三輪修正刪除；現行重點見下方展示層「重點 1–3」與「命題回覆」（數字以公式帶出）。"
FADE_LAB = {"scn_fade_g": "收斂法：各情境 FY2032 營收年增目標（延伸期收斂用）",
            "scn_fade_m": "收斂法：各情境中段淨利率目標（延伸期收斂用）"}


def _rows(ws, prefix, col=1):
    return [r for r in range(1, ws.max_row + 1) if str(ws.cell(r, col).value or "").startswith(prefix)]


def apply(wb, stub_note=None, stub_src=None, stub_lab=None, g_capped=False, obs_extra=None, readme_repl=()):
    """stub_note：Val「＋Stub 期 FCFF」列 E 欄說明（字串或公式）；stub_src：In 頁 stub_frac 列 I 欄說明；stub_lab：DCF 頁每股價值列 B 欄說明。
    g_capped：g 受 rf 上限（g_tv 高＝基準）。obs_extra：{README 列首字串: 新值或 None}（公司自訂的舊文處理）。readme_repl：[(舊, 新)] README 字串替換。"""
    log = []
    # 23：情境名稱
    n = 0
    for ws in wb:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str):
                    w = v
                    for a, b in SCN_REPL:
                        w = w.replace(a, b)
                    if w != v:
                        c.value = w; n += 1
    log.append(f"情境名稱 {n} 格")
    # 15：情境驅動對照
    IN = wb["In"]
    s0 = _rows(IN, "情境驅動對照")
    assert len(s0) == 1, s0
    r = s0[0] + 2
    while r <= IN.max_row and str(IN.cell(r, 1).value or "").startswith("scn_"):
        k, u = IN.cell(r, 1).value, IN.cell(r, 3).value
        if k in FADE_LAB:
            IN.cell(r, 2).value = FADE_LAB[k]
        if u in ("%", "pt"):
            for cc in (4, 5, 6):
                IN.cell(r, cc).number_format = "0.0%"
        elif u == "x":
            for cc in (4, 5, 6):
                IN.cell(r, cc).number_format = "0.00"
        r += 1
    # 24：上檔空間級距
    V = wb["Val"]
    for rr in _rows(V, "上檔空間級距"):
        if V.cell(rr, 2).value is not None:
            V.cell(rr, 1).value = "上檔空間級距（六級，供分組）"
            V.cell(rr, 5).value = "以「上檔空間：加權目標價」分六級（Andy 2026-10-25 14:00），只供分組；投資建議用下一列的五級固定規則"
    # 14：stub 標籤
    if stub_note is not None:
        rr = _rows(V, "＋Stub 期 FCFF")
        assert len(rr) == 1, rr
        V.cell(rr[0], 5).value = stub_note
    if stub_src is not None:
        rr = [x for x in range(1, IN.max_row + 1) if IN.cell(x, 1).value == "stub_frac"]
        assert len(rr) == 1, rr
        IN.cell(rr[0], 9).value = stub_src
    if stub_lab is not None and "DCF" in wb.sheetnames:
        D = wb["DCF"]; m = 0
        for x in range(1, D.max_row + 1):
            if str(D.cell(x, 2).value or "").startswith("stub FCFF＝"):
                D.cell(x, 2).value = stub_lab; m += 1
        assert m >= 3, m
    # 16：g 上限
    if g_capped:
        SE = wb["Sens"]
        rr = [x for x in range(1, SE.max_row + 1) if SE.cell(x, 1).value == "永續成長率 g ±1pt"]
        assert len(rr) == 1, rr
        SE.cell(rr[0], 1).value = "永續成長率 g −1pt／上限 rf"
    # 22：README 舊重點觀察
    RD = wb["README"]
    done2 = False
    for x in range(1, RD.max_row + 1):
        c = RD.cell(x, 1); v = c.value
        if not isinstance(v, str):
            continue
        if obs_extra:
            hit = [k for k in obs_extra if v.startswith(k)]
            if hit:
                c.value = obs_extra[hit[0]]; continue
        if v.startswith("=") and "重點觀察 1：" in v[:20] and '&"以下為 2026-10-29' in v:
            c.value = v[:v.index('&"以下為 2026-10-29')]
        elif v.startswith("重點觀察 1：") and "以下為 2026-10-29" in v:
            c.value = v[:v.index("以下為 2026-10-29")].rstrip()
        elif v.startswith(("重點觀察 2：", "重點觀察 3：")):
            c.value = None if done2 else OBS_NOTE
            done2 = True
        for a, b in readme_repl:
            if isinstance(c.value, str) and a in c.value:
                c.value = c.value.replace(a, b)
    log.append("README 舊重點觀察已清")
    return log
