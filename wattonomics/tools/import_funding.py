# 把 wattonomics/evidence_raw/F_funding.csv、P_facility_power.csv 整理成 builder/seed/funding_v0.1.1.json（FundRaw 頁、Facility 頁的來源），
# 附 G 段獨立核對判定（G_review_funding_facility.md，以它為準）與承重小誤更正。
# 用法：python3 wattonomics/tools/import_funding.py
# 列識別：F#n＝F_funding.csv 第 n 行；P#n＝P_facility_power.csv 第 n 行（第 1 行為表頭）。
# 採用（use＝1）：G 段判「屬實」、「小誤」且已依 G 段更正、或 G 段 §4 以 csv 值重算確認的列；「無法核對」與未抽查的列 use＝0（只列在 FundRaw）。
import csv, json, os, re

WT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RAW = os.path.join(WT, "evidence_raw")
OUT = os.path.join(WT, "builder", "seed", "funding_v0.1.4.json")


def rng(s):
    out = set()
    for part in s.split():
        m = re.match(r"(\d+)(?:–(\d+))?$", part)
        if m:
            a, b = int(m.group(1)), int(m.group(2) or m.group(1))
            out.update(range(a, b + 1))
    return out


# G 段 §1 判定（F 列）
F_TRUE = rng("44 45 46 47 52 58 56 69 70 71 72 74 75 77 80 81 87 89 90 91 92 93 94 95 96 97 100 101 103 104 105 106 107 116 117 118 120 121 122 124 126 127 "
             "128 129 130 131 132 134 136 7 8 9 13 14 15 19 20–22 27 29 30 31 137 138 140 141 142 144 145 146 147 148 150 154 155 158 159 173 177–178 183 187 190 192 "
             "205 206–208 210 211 212 213 214 217–218 220 222–225 227")
F_MINOR = rng("51 49 55 79 99 102 123 133 170 180 191 199 209")
F_UNVER = rng("50 57 16 151 152 162 167")
# G 段 §4 以 csv 值重算確認（未在 §1 逐列抽查）
F_REC = rng("10 11 17 186 39 40 41 42 43 64 65 66 67 68 73 98 164 165 168")
# 承重小誤更正（chat 端 v0.1.1 指示 2）與其他 §1 小誤更正
FIX = {
    191: dict(note="G 段更正（承重）：36 為初步洽談（Bloomberg 2026-08-04），未完成；『累計 >71』為 ZeroHedge 轉述招股書草稿、未證實。已完成只有 35（F190）", item="gpu_debt_talks"),
    133: dict(note="G 段更正（承重）：不可抗力為 Bloomberg 報導（Oracle 致 Blue Owl），通知日期不明；Oracle 稱 Jupiter「remains on schedule」", tag="[Interested-party]（記者報導）"),
    49: dict(note="G 段更正：依 8-K 票面利率 2066 6.500%、2046 6.250%、2036 5.450%（csv 原寫 6.283%、5.52% 可能為殖利率）",
             detail="USD 25B 優先無擔保；2066 6.500%、2046 6.250%、2036 5.450%（8-K 票面利率）"),
    55: dict(note="G 段更正：出處改為 2026Q2 法說（CFO；W.Media），新聞稿無資本支出指引", source_title="W.Media：Alphabet 2026 capex guidance（CFO，Q2 法說）",
             url="https://w.media/ai-infrastructure-demand-pushes-alphabets-2026-capex-guidance-to-us-205-billion/"),
    79: dict(note="G 段更正：出處改為 2026Q2 法說（MarketBeat 摘要），新聞稿無資本支出展望", source_title="MarketBeat：Amazon Q2 Earnings Call Highlights（Jassy）",
             url="https://www.marketbeat.com/instant-alerts/amazoncom-q2-earnings-call-highlights-2026-07-30/"),
    123: dict(note="G 段更正：出處改為 FY27Q1 法說（2026-09-10；MarketBeat「not exceeding $70 billion」），非新聞稿；引述非原文", source_title="Oracle Q1 FY27 法說（CFO）",
              url="https://www.fool.com/earnings/call-transcripts/2026/09/11/oracle-orcl-q1-2027-earnings-call-transcript/"),
    51: dict(note="G 段更正：刪除『美國史上最大股權發行』（8-K Ex.99.2、Bloomberg Law 皆無）"),
    199: dict(note="H 段更正（10-28，取代 G 段）：這是 Anthropic 付給 SpaceX 的算力採購，不是融資 → 不採用；『至 2029-05』原文有（Reuters 經 Yahoo），撤回 G 段刪除；年化 15 為 [Derived] 上限（爬坡期減免）",
              detail="付 SpaceX（xAI Colossus、Colossus II）每月 1.25 至 2029-05，年化 15 為上限；90 天可終止（SRC_ANT_227）"),
    99: dict(note="G 段更正：定價 4/30、交割 2026-05（10-Q）；Q 段第五節：10-Q『In May 2026, we issued』→ 期間改 2026-05（同落 9M2026 桶）", period="2026-05"),
    102: dict(note="G 段更正：債 12.5（非 12.3）；評等 A+／AA- 不在主來源"),
    170: dict(note="G 段更正：20–25 依 Zacks 2026-09-11；『約 60% 內部』是 2026-02 針對 16–20 的說法，套到 20–25 須註明"),
    180: dict(note="G 段更正：『a portion of the value』（非完工價值）；記者轉述、尚未簽約"),
    209: dict(note="G 段更正：用途『償還 xAI 舊債』不在 rexshares；與 F214 的 25 債為再融資關係，不重複計入"),
}
NB_6K = dict(source_title="Nebius 6-K 2026-08-12 附件 99.1／99.2（Q2 2026 財報新聞稿、未經查核簡明合併財報）",
             url="https://www.sec.gov/Archives/edgar/data/1513845/000110465926094568/tm2622968d1_ex99-1.htm",
             note="G 段更正：原 url 指向 2025 年度 20-F（不含 H1）；改引 6-K 2026-08-12（nebius/data/nebius_facts_20261007.json *.q2_26）")
for n in range(162, 170):
    FIX[n] = dict(NB_6K)

# Q 段第五節（2026-10-31，Amazon、Meta 2026Q3 10-Q 獨立核對；evidence_raw/Q_review_2026q3.md 第五節、Q_add_2026q3_10q.md）：既有列更正與新列標記更正
# 值的更正（value）與期間更正（period）沿用 FIX 機制（v0.1.7 起 FIX 支援 value、period 兩鍵）；doc_date 只存在 CSV，Meta 10-Q 各列 10-29 → 10-21 直接改 CSV
Q5 = "Q 段第五節更正（10-Q）"
FIX.update({
    74: dict(note=Q5, detail="USD 37＋EUR 14.5，19 檔；史上最大公司債；10-Q 附註 5：March 2026 Notes 37.0＋Euro Notes €14.5B（9-30 帳面 16.427）"),
    75: dict(note=Q5 + "：保留 7 月報導的歸屬，並註明與 10-Q 不符",
             detail="8 檔、2029–2066、4.20%–6.25%（含 2029 浮動 0.75；10-Q 附註 5）；7 月報導稱公司表示 2026 年不再發債（原句待回 TNW／Bloomberg 確認範圍，可能只指美元債），"
                    "但 10-Q 顯示其後 9 月仍發英鎊債 £4.3B（帳面 5.637）並動用定期貸款 17.5，MD&A 稱 2026 年將再融資（We expect to undertake additional financing activities in 2026）；已含於 F#253"),
    80: dict(note=Q5, detail="對 OpenAI 累計 50（Q1 15＋Q2 13.7＋Q3 21.3；10-Q MD&A：including $21.3 billion invested in Q3 2026）；OpenAI 承諾 AWS 38＋100"),
    81: dict(note=Q5 + "（附帶發現）：原 18.0 含 2023–2025 年可轉債 8.0，不屬 2026 年；10-Q 載 2026 年實際投入 10.0 → 值改 10.0", value=10.0,
             detail="對 Anthropic 2026 年投入 10.0（Series G 5＋Series H 5；10-Q）；累計 18 另含 2023–2025 年可轉債 8（不屬 2026，不計）；剩餘額度 15 依算力交付里程碑；可轉債公允價值 97.9"),
    100: dict(note=Q5 + "：原 22.178 只取非流動部分（9-30 有 2.747 重分類至流動），值改面額口徑 25.000（chat 端決定；與 F#263、Alphabet F#50 一致）", value=25.0,
              detail="面額口徑：債務面額 59.0 → 84.0，增加 25.000（10-Q 附註 9；5 月發債 25.0，淨收入 24.910）；帳面 58.744 → 83.669（非流動 80.922＋一年內到期 2.747，後者列應計費用及其他流動負債），"
                     "帳面口徑增加 24.925；現金與有價證券 81.592 → 82.449 → 9-30 淨負債約 1.2（原寫『淨現金約 0』漏計一年內到期 2.747）"),
    252: dict(note=Q5 + "：補 10-Q 第二來源，標記改 [Verified]", tag="[Verified]",
              detail="現金流量表無買回、股利列（2026 未買回、無股利）；10-Q 附註 6：There were no repurchases of our common stock during the nine months ended September 30, 2025 or 2026"),
    253: dict(note=Q5 + "：detail 改寫（『＝』改『≈』）",
              detail="長期借款收入 115.1 ≈ 票據 97.3（3 月 USD 37.0、EUR 14.5≈16.4；5 月 CHF 2.8≈3.4；6 月 CAD 14.0≈9.8；7 月 USD 25.0；9 月 GBP 4.3≈5.6；外幣以 9-30 帳面，F#259）＋9 月動用定期貸款 17.5（F#260）；"
                     "差 0.3 為外幣匯率與折價；Q3 單季 48.102；還款 2.759（2021 年票據到期）；短期借款收入 20.693、還款 19.609（另列不計）；9-30 長期借款帳面 173.125（另一年內到期 6.827）"),
    256: dict(note=Q5 + "：補 10-Q 第二來源",
              detail="指引：2027 現金資本支出將 meaningfully more（較 2026 約 220 明顯增加，未給數字；> 220 為推論下限，不入數值）；10-Q MD&A：We expect cash capital expenditures to meaningfully increase in 2027；"
                     "另稱 AWS 2027、2028 各年營收成長可望 40% 以上；對 2027、2028 及部分 2029 的大塊容量（數百 MW，有時達 1 GW）能見度很好"),
    257: dict(note=Q5 + "：detail 三處改寫；doc_date 留空",
              detail="9-30 債務總額（存量）83.7（CFO 口徑；10-Q 帳面 83.669＝非流動 80.922＋一年內到期 2.747（2027-08 到期 2.75，列應計費用及其他流動負債），見 F#272）；"
                     "現金與有價證券 82.4 → 淨負債約 1.3（10-Q：83.669 − 82.449＝1.220；F#100 原『淨現金約 1.1』漏計一年內到期部分，已更正）；Q3 單季含融資租賃本金資本支出 34.9、FCF −2.1；"
                     "宣布新設無表決權股（股票股利 1:1，預計 2027-01 下旬交易），用途含募資、員工薪酬、併購；通話當時稱「目前無具體融資計畫」，其後 10-26 FWP 即定價 37.5（F#106、F#107）；"
                     "通話日期逐字稿與 10-Q 都未載；10-Q 於 2026-10-21 簽署，附註 14 期後事項已載 10 月提交 Restated Charter 新設 Class C 並宣告 1:1 股票股利（生效日 2027-01-19），且未提 10-26 的遠期出售與特別股 → 新股類別宣布不晚於 10-21，法說推定不晚於 10-21；2027 資本支出指引待 Q4 法說"),
    270: dict(note=Q5 + "：值為兩數相加 → [Derived]", tag="[Derived]"),
    271: dict(note=Q5 + "：值為兩數相加 → [Derived]", tag="[Derived]"),
    274: dict(note=Q5 + "：值為兩數相加 → [Derived]", tag="[Derived]"),
    275: dict(note=Q5 + "：值為兩數相加 → [Derived]", tag="[Derived]"),
})

# H 段（10-28 API 旗標人工複核，evidence_raw/H_review_api_flags.md）：採購／投資存量不是融資 → 不採用
H_DROP = {
    31: "H 段更正：Microsoft 租用 neocloud 算力是採購承諾（Microsoft 為付款方），不是資金來源 → 不採用",
    199: None,   # 見 FIX[199]
    222: "H 段連帶更正（同 F#226 口徑）：NVIDIA 非上市權益存量是投資組合，不是供應商融資流量 → item 改 investment_stock_nonmkt、不採用",
}
H_ITEM = {222: "investment_stock_nonmkt"}

# Q 段（v0.1.3–v0.1.4，2026Q3／FY27Q1 財報）：F#228–246 由 Andy 下載的（F#247–254 Amazon 由公司新聞稿） 10-Q、新聞稿、法說逐字稿取數，經未參與代理核對（evidence_raw/Q_review_2026q3.md）
Q_TRUE = rng("228–246 247–254 255–257 258–275")   # 258–275：Amazon、Meta 2026Q3 10-Q（Q_review 第五節：屬實 14、小誤 4 已改 tag）   # 247–254：Amazon 2026Q3 新聞稿（Q_review 第三節：屬實 7、小誤 1 已改 tag）；255–257：Amazon、Meta 逐字稿（第四節：屬實 1、小誤 2 已改）
# 被 Q 段取代（同期間較新或實收口徑）：Alphabet 2026H1 各列與 8 月發債（併入 9M2026 現金流量表）、F#51 股權公告 84.75（含未動用 ATM，改用實收 61.933）、
# F#55 2026 指引（Q2 法說，Q3 重申另列）；Microsoft F#29 曆年 175（改 170）、F#30 FY27Q1 指引（已有實績）
Q_SUPERSEDED = {44: "9M2026", 45: "9M2026", 46: "9M2026", 47: "9M2026", 48: "9M2026", 49: "9M2026（8 月發債已在 9M 現金流量表內）",
                51: "9M2026 實收 61.933（F#244）", 55: "Q3 法說重申（F#245）", 29: "Q1 法說 170（F#237）", 30: "FY27Q1 實績（F#228–234）",
                # Amazon（Q_review 第三節）：H1 各列併入 9M2026；3 月、7 月發債已在 9M 長期借款收入 115.1（F#253）內，保留作事件紀錄不計入；指引改 Q3 重申
                69: "9M2026（F#247）", 70: "9M2026（F#248）", 71: "9M2026（F#249）", 72: "9M2026（F#251）",
                74: "9M2026 長期借款收入 115.1 已含（F#253）", 75: "9M2026 長期借款收入 115.1 已含（F#253）", 79: "Q3 新聞稿重申（F#254）",
                # Q_review 第五節：Amazon 10-Q 9-30 期末存量取代 6-30（F#77 必要：lnc 同桶會加總）
                77: "F#258（9-30）", 76: "F#263", 78: "F#266"}


def h_notes():
    """H_corrections.csv 的 F 列更正理由（併入 gnote）"""
    out = {}
    p = os.path.join(RAW, "H_corrections.csv")
    if os.path.exists(p):
        for c in csv.DictReader(open(p, encoding="utf-8-sig")):
            if c["file"] == "F_funding.csv" and c["field"] not in ("(none)", "(note)"):
                out.setdefault(int(c["row"]), []).append(f"{c['field']}：{c['reason']}")
    return out


# 期間桶：每位玩家 2024／2025／2026（最新期）＝(標籤, 起, 迄)；其餘 post（期後）或 guide（指引）
BUCKET = {
    "Microsoft": {2024: ("FY2024", "2023-07", "2024-06"), 2025: ("FY2025", "2024-07", "2025-06"), 2026: ("FY2026", "2025-07", "2026-06")},
    "Alphabet": {2024: ("CY2024", "2024-01", "2024-12"), 2025: ("CY2025", "2025-01", "2025-12"), 2026: ("9M2026", "2026-01", "2026-09")},   # v0.1.3：2026H1 → 9M2026
    "Amazon": {2024: ("CY2024", "2024-01", "2024-12"), 2025: ("CY2025", "2025-01", "2025-12"), 2026: ("9M2026", "2026-01", "2026-09")},   # v0.1.3：2026H1 → 9M2026
    "Meta": {2024: ("CY2024", "2024-01", "2024-12"), 2025: ("CY2025", "2025-01", "2025-12"), 2026: ("9M2026", "2026-01", "2026-09")},
    "Oracle": {2024: ("FY2024", "2023-06", "2024-05"), 2025: ("FY2025", "2024-06", "2025-05"), 2026: ("FY2026", "2025-06", "2026-05")},
    "CoreWeave": {2024: ("CY2024", "2024-01", "2024-12"), 2025: ("CY2025", "2025-01", "2025-12"), 2026: ("2026H1", "2026-01", "2026-06")},
    "Nebius": {2025: ("CY2025", "2025-01", "2025-12"), 2026: ("2026H1", "2026-01", "2026-06")},
    "xAI（SpaceX）": {2025: ("CY2025", "2025-01", "2025-12"), 2026: ("2026H1", "2026-01", "2026-06")},
    "OpenAI": {2025: ("CY2025", "2025-01", "2025-12"), 2026: ("2026 迄今", "2026-01", "2026-10")},
    "Anthropic": {2025: ("CY2025", "2025-01", "2025-12"), 2026: ("2026 迄今", "2026-01", "2026-10")},
}
PKEY = {"Microsoft": "msft", "Alphabet": "googl", "Amazon": "amzn", "Meta": "meta", "Oracle": "orcl", "CoreWeave": "crwv", "Nebius": "nbis",
        "xAI（SpaceX）": "xai", "OpenAI": "oai", "Anthropic": "ant"}
STOCK = {"lease_not_commenced", "net_debt_change", "op_lease"}


def ym(period):
    """期間字串 → (起 YYYY-MM, 迄 YYYY-MM) 或 None"""
    p = period.strip()
    m = re.match(r"(FY|CY)(\d{4})$", p)
    if m: return ("FY", int(m.group(2))) if m.group(1) == "FY" else (f"{m.group(2)}-01", f"{m.group(2)}-12")
    m = re.match(r"(\d{4})H1$", p)
    if m: return (f"{m.group(1)}-01", f"{m.group(1)}-06")
    m = re.match(r"9M(\d{4})$", p)
    if m: return (f"{m.group(1)}-01", f"{m.group(1)}-09")
    m = re.match(r"(\d{4})Q([1-4])$", p)
    if m:
        q = int(m.group(2)); return (f"{m.group(1)}-{3 * q - 2:02d}", f"{m.group(1)}-{3 * q:02d}")
    m = re.match(r"(\d{4})-Q([1-4])$", p)
    if m:
        q = int(m.group(2)); return (f"{m.group(1)}-{3 * q - 2:02d}", f"{m.group(1)}-{3 * q:02d}")
    m = re.match(r"(\d{4})-(\d{2})", p)
    if m: return (f"{m.group(1)}-{m.group(2)}", f"{m.group(1)}-{m.group(2)}")
    return None


def bucket(player, period):
    bk = BUCKET.get(player)
    if not bk: return "other"
    for b, (lab, s, e) in bk.items():
        if period == lab: return b
    if period.startswith("CY2026") or period.startswith("FY2027") or period == "FY27Q1" or period.startswith("2023–"):
        return "guide" if period.startswith(("CY2026", "FY2027")) else "post"
    r = ym(period)
    if not r or r[0] == "FY": return "post"
    s0, e0 = r
    for b, (lab, s, e) in bk.items():
        if s <= s0 and e0 <= e: return b
    return "post"


def main():
    rows = []
    with open(os.path.join(RAW, "F_funding.csv"), encoding="utf-8-sig") as f:
        rd = list(csv.reader(f))
    h = rd[0]
    HN = h_notes()
    for n, row in enumerate(rd[1:], 2):
        d = dict(zip(h, row))
        if n in F_TRUE: v = "屬實"
        elif n in F_MINOR: v = "小誤（已更正）"
        elif n in F_UNVER: v = "無法核對"
        elif n in F_REC: v = "G§4 重算確認"
        else: v = "未抽查"
        fx = FIX.get(n, {})
        if n in range(162, 170) and v in ("無法核對", "未抽查"):
            v = "已更正（改引 6-K）"
        item = H_ITEM.get(n, fx.get("item", d["item"])).split("（")[0].strip()   # H 段 item 帶括號說明，比對用前段
        if n == 205: item = "capex_ai_seg"     # xAI 段資本支出（合併數另列 F207），避免與合併資本支出重複
        if n == 209: item = "bridge_refi"       # 過橋 20：再融資，與 F214 債 25 不重複（G 段 §4）
        # v0.1.4：F#257 Meta 債務總額為存量 debt_level（不在 FITEMS，不進淨負債變化加總；只列 FundRaw）
        if n == 103: item = "jv_exposure"       # 合資最大損失曝險（存量），不與當期表外出資相加
        if n == 58: item = "jv_exposure"        # v0.1.3：Google 對第三方站點租約擔保最大曝險（或有、存量），不是 Alphabet 的資金來源（9M2026 期間後會落入 2026 桶）
        if n in (28, 54, 105): item = "commitments"   # 採購／合約承諾（存量），不是融資
        if n == 204: item = "valuation"         # 併購估值，不是募資
        if item == "net_debt_change" and "水位" in d["detail"]: item = "net_debt_level"
        if n in Q_TRUE:
            v = "Q 段核對"
        if n in Q_SUPERSEDED:
            v = "由 Q 段取代：" + Q_SUPERSEDED[n]
        use = 1 if v in ("屬實", "小誤（已更正）", "G§4 重算確認", "已更正（改引 6-K）", "Q 段核對") else 0
        gnote = fx.get("note", "")
        if n in H_DROP:
            use = 0; v = "H 段改不採用"
            if H_DROP[n]: gnote = (gnote + "；" if gnote else "") + H_DROP[n]
        elif n in HN:
            gnote = (gnote + "；" if gnote else "") + "H 段更正（10-28）：" + "；".join(HN[n])
        try: val = float(d["value"])
        except ValueError: val = None
        if "value" in fx: val = fx["value"]          # v0.1.7：Q 段第五節值的更正（F#81、F#100）
        period = fx.get("period", d["period"])       # v0.1.7：期間更正（F#99）
        player = d["player"]
        rows.append(dict(id=f"F#{n}", player=player, pkey=PKEY.get(player, ""), period=period, bucket=str(bucket(player, period)),
                         item=item, value=val, unit=d["unit"], detail=fx.get("detail", d["detail"]), tag=fx.get("tag", d["tag"]),
                         source=fx.get("source_title", d["source_title"]), url=fx.get("url", d["url"]), verdict=v, gnote=gnote, use=use))
    # xAI 2025：S-1 新增債務 16（F#205 detail；G 段 F205 屬實）建成 GPU 債項目，標記照原列
    f205 = next(e for e in rows if e["id"] == "F#205")
    rows.append(dict(f205, id="F#205b", item="gpu_debt", value=16.0, detail="xAI 段 2025 新增債務 16（GPU）；由 F#205 detail 拆出",
                     gnote="由 F#205（G 段屬實）拆出；F#200–202 未抽查維持不採用"))
    # 期初存量（G 段 §1 正文所載，供『尚未起租變化』）：Microsoft 2025-06 為 92.7（F27）；Oracle FY26 末 260（F129）；Amazon 2025 年底 96.4（F77 detail）
    for pk, player, period, b, val, note in (("msft", "Microsoft", "2025-06-30", "2025", 92.7, "G 段 F27：前一年 6 月為 92.7（btw.media 引 10-K）"),
                                              ("orcl", "Oracle", "2026-05-31", "2026", 260.0, "G 段 F129：FY26 末 260（Reuters）"),
                                              ("amzn", "Amazon", "2025-12-31", "2025", 96.4, "F77 detail：2025 年底 96.4（G 段 F77 屬實）")):
        rows.append(dict(id=f"G:{pk}_lnc0", player=player, pkey=pk, period=period, bucket=b, item="lease_not_commenced", value=val, unit="USD B",
                         detail="尚未起租租賃（期初存量）", tag="[Verified]" if pk != "orcl" else "[Interested-party]",
                         source="G_review_funding_facility.md §1", url="", verdict="屬實", gnote=note, use=1))
    prow = []
    with open(os.path.join(RAW, "P_facility_power.csv"), encoding="utf-8-sig") as f:
        rd = list(csv.reader(f))
    h = rd[0]
    P_TRUE = rng("3 4 5 6–9 10 11 12–13 14 15 19–20 21 22 23 24 27 29 31 38 39 40 41 44 49 54 55 56 57 58 59 63")
    P_MINOR = rng("2 16–18 28 30 42 62 50")
    for n, row in enumerate(rd[1:], 2):
        d = dict(zip(h, row))
        v = "屬實" if n in P_TRUE else ("小誤（已更正）" if n in P_MINOR else "未抽查")
        prow.append(dict(id=f"P#{n}", item=d["item"], sub=d["sub_item"], metric=d["metric"], low=d["low"], base=d["base"], high=d["high"], unit=d["unit"],
                         basis=d["basis"], as_of=d["as_of"], tag=d["tag"], source=d["source_title"], url=d["url"], verdict=v))
    json.dump(dict(funding=rows, facility=prow, bucket={k: {str(b): v for b, v in bk.items()} for k, bk in BUCKET.items()}),
              open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=0)
    from collections import Counter
    print(f"{OUT}：F {len(rows)} 列 {dict(Counter(r['verdict'] for r in rows))}；採用 {sum(r['use'] for r in rows)}；P {len(prow)} 列")
    print(Counter((r["pkey"], r["bucket"]) for r in rows if r["use"]))


if __name__ == "__main__":
    main()
