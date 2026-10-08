"""來源四欄（等級、立場、立場說明、一手／二手）的機械初評規則（Anthropic 版；改寫自 OpenAI v0.6 V10）。

依據 data/anthropic_src.yaml 每列的 tag（標記）、party（利害方與誘因）與 source（來源文字）判定；標「CC 初評」。
規則是分類邏輯，不含任何數值參數；無法判定者填「未評」。
"""
UNRATED = "未評"
RELAY = ["經 ", "轉述", "轉載", "引述", "引 ", "摘要"]                      # 媒體轉述他人
LEAK = ["招股書", "說明書", "外流", "內部文件", "機密", "投資人"]         # 外流文件或內部資料
OFFICIAL = ["anthropic.com", "Anthropic〈", "新聞稿", "10-Q", "10-K", "8-K", "S-1", "IPO 申報", "定價頁", "部落格", "公告"]
MEDIA = ["Reuters", "Bloomberg", "WSJ", "FT", "Fortune", "CNBC", "Axios", "TechCrunch", "The Information", "Forbes", "New York Times", "NYT",
         "PYMNTS", "Khaleej", "KuCoin", "Yahoo", "SaaStr", "Sacra", "The Register", "DCD", "Data Center", "Newsquawk", "Mobile World", "Business"]


def classify(row: dict) -> dict:
    src = str(row.get("source") or "")
    tag = str(row.get("tag") or "")
    party = row.get("party")
    if tag == "Derived":
        hand = "推導"
    elif any(k in src for k in LEAK) and any(k in src for k in RELAY + MEDIA):
        hand = "二手（含一手轉載）"            # 媒體轉述外流文件（與 OpenAI F2 同義）
    elif any(k in src for k in OFFICIAL) and any(k in src for k in RELAY):
        hand = "二手（含一手公告轉載）"
    elif any(k in src for k in OFFICIAL):
        hand = "一手"
    elif any(k in src for k in MEDIA + RELAY):
        hand = "二手"
    else:
        hand = UNRATED
    if tag.startswith("Interested-party") or (tag == "Verified" and party):
        stance, why = "利害關係方", str(party or UNRATED)
    elif tag == "Analogy":
        stance, why = "第三方估計", str(party) if party else "第三方估計；非公司揭露，作類比或區間"
    elif tag == "Verified":
        stance, why = "中立或已核對", "公開文件可核對（Verified）"
    elif tag == "Derived":
        stance, why = "推導", "由其他 SRC_ANT 列推得（本活頁簿以公式表示）"
    else:
        stance, why = UNRATED, UNRATED
    if "SemiAnalysis" in src or "InferenceX" in src:
        why = "SemiAnalysis：偏樂觀，須第二獨立來源"
    grade = {"一手": "1"}.get(hand, "2" if hand.startswith("二手") else UNRATED)
    return dict(grade=grade, stance=stance, why=why, hand=hand)
