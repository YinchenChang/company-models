"""來源四欄（來源等級、立場、立場說明、一手／二手）的機械初評規則（沿用 OpenAI v0.6 V10 精神；智譜 Z2 改寫）。

智譜的 data/zhipu_src.yaml 每列已有標記（tag）與利害方（party），本模組只依這兩欄與出處、網址做分類，
不含任何數值參數。規則無法判定者填「未評」，欄 X 標「CC 初評」。

- 一手：港交所披露易（HKEXnews）招股章程、年度／中期業績公告、配售與併購公告；公司 IR 網站轉載的同一份公告；
  公司官方定價頁（bigmodel.cn、z.ai、docs.z.ai）；App Store 官方頁。
- 二手：媒體、券商、第三方整理站（含轉錄官方價目表者）。
- 來源等級：1＝一手已讀原文；2＝二手已讀或轉載（含 Interested-party 與 Analogy）；3＝待查核（basis／note 含「未核得」「待查」）。
"""
OFFICIAL_URL = ["hkexnews.hk", "eurolandir.com", "bigmodel.cn", "z.ai", "zhipuai.cn", "apps.apple.com"]
OFFICIAL_SRC = ["HKEXnews", "披露易", "招股章程", "年度業績公告", "中期業績公告", "年報", "中期報告", "公告", "官方定價頁", "定價頁", "bigmodel.cn", "z.ai"]
MEDIA = ["媒體", "轉述", "報導", "引述", "知情人士", "網易", "同花順", "證券日報", "36氪", "華爾街見聞", "每日經濟", "鈦媒體", "澎湃", "鉅亨",
         "財聯社", "虎嗅", "IT之家", "AIbase", "Bloomberg", "Reuters", "SCMP", "量子位", "券商", "證券", "研報", "奇連", "CodePick", "Digital Applied",
         "UsagePricing", "aipricing", "Simply Wall", "Investing", "上海有色", "SMM", "AtomGit", "CSDN", "GIGAZINE", "Tom's Hardware", "The Register",
         "Unite.AI", "Lenovo", "DatabaseMart", "QuestMobile", "OpenRouter", "Wikipedia", "Labmemo", "鳳凰網", "騰訊新聞", "新浪", "格隆匯", "東方財富"]
UNRATED = "未評"
PENDING_MARKS = ["未核得", "待查", "待核"]
VERIFIED_WHY = "經審計或經審閱的公司公告（誘因：上市後股價、配售；數字可由公告原文核對）"
PRICE_WHY = "公司公開且具約束力的報價，誤報誘因低"
ANALOGY_WHY = "第三方量測或類比資料，非智譜利害方；只作區間參考"
DERIVED_WHY = "由其他 SRC 列推算（見備註）；立場同其來源"


def classify(row: dict) -> dict:
    """回傳 grade／stance／why／hand。row 為 zhipu_src.yaml 的一列。"""
    src = f"{row.get('source') or ''} {row.get('url') or ''}"
    note = f"{row.get('basis') or ''} {row.get('note') or ''}"
    tag = str(row.get("tag") or "")
    party = row.get("party")
    official = any(k in (row.get("url") or "") for k in OFFICIAL_URL) or any(k in (row.get("source") or "") for k in OFFICIAL_SRC)
    media = any(k in src for k in MEDIA)
    if official and media:
        hand = "二手（含一手公告轉載）"
    elif official:
        hand = "一手"
    elif media:
        hand = "二手"
    elif tag == "Derived":
        hand = "推算"
    else:
        hand = UNRATED
    if any(k in note for k in PENDING_MARKS):
        grade = "3"
    elif hand == "一手" and tag in ("Verified", "Interested-party"):
        grade = "1"
    elif tag in ("Verified", "Interested-party", "Analogy", "Derived"):
        grade = "2"
    else:
        grade = UNRATED
    if tag == "Interested-party":
        stance, why = "利害關係方", (str(party) if party else UNRATED)
    elif tag == "Verified" and hand == "一手" and any(k in (row.get("url") or "") for k in ("bigmodel.cn", "z.ai", "apps.apple.com")):
        stance, why = "利害關係方", PRICE_WHY
    elif tag == "Verified":
        stance, why = "利害關係方", VERIFIED_WHY if hand.startswith("一手") or "公告" in hand else "公司數字經第三方轉錄（見備註）；誘因同公司公告"
    elif tag == "Analogy":
        stance, why = "第三方", ANALOGY_WHY
    elif tag == "Derived":
        stance, why = "推算", DERIVED_WHY
    else:
        stance, why = UNRATED, UNRATED
    if "SemiAnalysis" in src or "SemiAnalysis" in note:
        why = "SemiAnalysis：偏樂觀，須第二獨立來源（不得單獨引用）"
    return dict(grade=grade, stance=stance, why=why, hand=hand)
