"""V10：來源四欄（等級、立場、立場說明、一手／二手）的機械初評規則（工作單 r2 P1-2）。

標「CC 初評」；規則無法判定者填「未評」並計入報告。規則是分類邏輯，不含任何數值參數。
"""
MEDIA = ["Reuters", "The Information", "WSJ", "FT/", "FT ", "Fortune", "Bloomberg", "CNBC", "Axios", "TNW", "TechCrunch", "VentureBeat",
         "AI Magazine", "GuruFocus", "Techzine", "MarkTechPost", "Epoch", "外流", "報導", "轉述", "引述", "知情人士", "New York Times", "NYT"]
OFFICIAL = ["公告", "揭露", "10-Q", "8-K", "10-K", "定價頁", "部落格", "DevDay", "證詞", "法說", "OpenAI 2026-02", "OpenAI（"]
COUNTERPARTY = ("Oracle", "AWS", "CoreWeave", "Cerebras", "Microsoft", "Nvidia")
INTERNAL_MARKS = ["簡報", "內部", "預測", "投資人", "股東", "外流", "財報", "證詞", "目標", "計畫", "S-1", "管理層"]
UNRATED = "未評"
PRICE_MARKS = ["pricing", "定價頁", "價格頁", "openai.com"]            # OpenAI 自己公布的牌價（V14）
PRICE_WHY = "公開且具約束力的報價，誤報誘因低"
CONTRACT_WHY = "合約雙方公告；OpenAI 有募資與算力規模訊號誘因，供應商有營收與股價訊號誘因；金額為承諾總額，非實付"   # 補充意見 G1


def trace_sources(rows: list) -> list:
    """H2：把出處為「同上」的列追溯為實際出處文字（沿前一列逐步展開）；追不到時保留原文。回傳被追溯的列清單。"""
    traced = []
    for i, r in enumerate(rows):
        src = r["source"] or ""
        r["source_base"] = src
        r["trace_chain"] = []
        if src.startswith("同上") and i:
            prev = rows[i - 1]
            r["source_raw"] = src
            r["source_base"] = prev["source_base"]
            r["trace_chain"] = [prev["metric"]] + prev["trace_chain"]
            chain = "←".join(f"『{m}』" for m in r["trace_chain"])
            r["source"] = f"{r['source_base']}（追溯自「同上」鏈：本列←{chain}；原文：{src}）"
            traced.append(r["metric"])
    return traced


def classify(row: dict, prev_source: str = "") -> dict:
    if row.get("cl"):                       # 列層級覆寫（H1、H4）：由 v05_map 明列
        return dict(row["cl"])
    src = row["source"] or ""
    if src.startswith("同上") or src.startswith("同"):
        src = prev_source + "；" + src
    if any(k in src for k in PRICE_MARKS):         # V14：OpenAI 自己公布的牌價列＝一手、利害關係方、誘因低
        chk0 = str(row.get("chk", ""))
        return dict(grade="3" if chk0.startswith("待查核") else "2", stance="利害關係方", why=PRICE_WHY, hand="一手")
    media = any(k in src for k in MEDIA)
    official = any(k in src for k in OFFICIAL)
    if "外流" in src:                              # F2：外流財報＝二手（含一手轉載）
        hand = "二手（含一手轉載）"
    elif media and official:
        hand = "二手（含一手公告轉載）"
    elif media:
        hand = "二手"
    elif official:
        hand = "一手"
    else:
        hand = UNRATED
    tag = str(row["tag"])
    text = " ".join(str(row.get(k, "")) for k in ("metric", "scope", "use")) + " " + src
    stance, why = UNRATED, UNRATED
    if tag.startswith("Interested-party"):
        stance = "利害關係方"
        party = any(c in src for c in COUNTERPARTY + ("OpenAI",))
        if any(k in text for k in INTERNAL_MARKS) or "融資" in text or "現金" in text:       # 先判內部／外流數據，再判交易對手
            why = "管理層、內部文件或外流數據：誘因為募資與議價，傾向呈現成長與算力需求"
        elif row["subject"] in COUNTERPARTY or any(c in src for c in COUNTERPARTY if c != "Nvidia" or "Nvidia" in row["metric"]):
            why = "交易對手公告或轉述：誘因為宣傳合約規模與客戶關係"
        elif official and party:
            why = "公司公告：誘因為宣傳成長與用戶規模"
        else:
            why = UNRATED
    elif tag == "Verified" and not media and (official or any(c in src for c in COUNTERPARTY + ("OpenAI",))):
        # G1：v0.5 標 Verified 的合約列，來源為 OpenAI 或交易對手的官方公告／文件 → 一手、利害關係方；非合約列（subject＝OpenAI）用一般說明
        stance = "利害關係方"
        why = CONTRACT_WHY if row["subject"] in COUNTERPARTY else "公司或交易對手公告：誘因為宣傳；內容可由公開文件核對（v0.5 標 Verified）"
        if hand == UNRATED:
            hand = "一手"
    if "SemiAnalysis" in src:
        why = "偏樂觀，須第二獨立來源（本列無第二來源）"
    chk = str(row.get("chk", ""))
    if chk.startswith("待查核"):
        grade = "3"
    elif tag.startswith(("Interested-party", "Verified")):
        grade = "2"
    else:
        grade = UNRATED
    return dict(grade=grade, stance=stance, why=why, hand=hand)
