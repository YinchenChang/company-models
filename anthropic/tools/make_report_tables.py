#!/usr/bin/env python3
"""由 builder/v05_map.py 產生報告附表（不手寫）：leafmap.csv、tagged_objects.csv、v9_splits.csv、v10_stats.md。
用法：python3 tools/make_report_tables.py [--out docs/reports] [--prefix 20261004_v0.6-P1]"""
import argparse
import csv
import json
import sys
from collections import Counter
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "builder"))
import source_rules  # noqa: E402
from v05_map import META_KEYS, SEP, build_map  # noqa: E402

KL = {"SRC": "SRC_OAI", "INP": "Inputs", "FORMULA": "公式（後續工作包）", "DUP": "重複併入", "SKIP": "不遷入", "TK": "不遷入；改取 TK_Link"}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", type=Path, default=REPO / "docs" / "reports")
    ap.add_argument("--prefix", default="20261004_v0.6-P1")
    a = ap.parse_args()
    R = build_map()
    with open(a.out / f"{a.prefix}_leafmap.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["v0.5 路徑", "v0.5 值", "去處類別", "去處", "理由／備註"])
        for p, v in R.leaves.items():
            k, ref, why = R.dest[p]
            w.writerow([p.replace(SEP, "."), json.dumps(v, ensure_ascii=False), KL[k], ref, why])
    rows = []

    def walk(o, path):
        if isinstance(o, dict):
            if "tag" in o:
                leaves = R.under(SEP.join(path))
                core = [p for p in leaves if p.split(SEP)[-1] not in META_KEYS] or leaves
                ds = sorted({(R.dest[p][0], R.dest[p][1]) for p in core})
                rows.append((".".join(path), o["tag"], len(leaves), "；".join(f"{KL[k]}：{r}" for k, r in ds)))
            for k, v in o.items():
                walk(v, path + [k])
        elif isinstance(o, list):
            for i, v in enumerate(o):
                walk(v, path + [str(i)])

    walk(R.data, [])
    with open(a.out / f"{a.prefix}_tagged_objects.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["v0.5 物件路徑", "v0.5 標記", "葉節點數", "去處（可多筆）"])
        w.writerows(rows)
    with open(a.out / f"{a.prefix}_v9_splits.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["多年期項目", "處理", "結果去處", "掃描後未拆"])
        for r in R.v9:
            w.writerow([r["item"], r["detail"], r["dest"], "是" if r["scanned_only"] else "否（已拆）"])
    cnt = {k: Counter() for k in ("grade", "stance", "hand", "why")}
    unrated = {k: [] for k in cnt}
    for i, r in enumerate(R.src_rows):
        c = source_rules.classify(r, R.src_rows[i - 1]["source"] if i else "")
        for k in cnt:
            cnt[k][c[k]] += 1
            if c[k] == source_rules.UNRATED:
                unrated[k].append(f"{r['id']} {r['metric']}")
    lines = [f"SRC_OAI {len(R.src_rows)} 列 V10 初評統計"]
    for k, lab in (("grade", "來源等級"), ("stance", "立場"), ("hand", "一手／二手"), ("why", "立場說明")):
        lines.append(f"- {lab}：" + "；".join(f"{v}＝{n}" for v, n in cnt[k].most_common()))
    for k, lab in (("grade", "來源等級"), ("hand", "一手／二手"), ("stance", "立場")):
        lines.append(f"- 『未評』明細（{lab}，{len(unrated[k])} 列）：" + "；".join(unrated[k]))
    (a.out / f"{a.prefix}_v10_stats.md").write_text("\n".join(lines) + "\n", encoding="utf-8")

    # ── E7 帳目 ────────────────────────────────────────────
    # (a) SRC_OAI／Inputs 列數變動逐筆對帳（初審 SHA ebe08a4 → 第二版 SHA 711cfec → 本版）
    recon = [
        ("SRC_OAI", "SRC_OAI_086", "訓練算力支出：2025（12.0）", "ebe08a4 有；711cfec 移除", "V6（r2 §5）：12.0＝10.59＋其他雲端 1.41，改為公式；ID 退役"),
        ("SRC_OAI", "SRC_OAI_092", "ChatGPT 週活躍用戶（WAU）：2026-02", "711cfec 新增", "V9 free：出處文字內的觀測值"),
        ("SRC_OAI", "SRC_OAI_093", "Go 付費用戶：2025 年底", "711cfec 新增", "V9 掃描 go：錨點"),
        ("SRC_OAI", "SRC_OAI_094", "ChatGPT Plus＋Pro 付費訂閱：2025-07", "711cfec 新增", "V9 plus：觀測值（公式依據）"),
        ("SRC_OAI", "SRC_OAI_095", "ChatGPT 付費訂閱：2025 年（內部文件）", "711cfec 新增", "V9 plus：觀測值"),
        ("SRC_OAI", "SRC_OAI_096", "Pro 用戶 FY 平均：2026（內部預測，1.0M）", "711cfec 新增；本版移除", "V12（r3）：SRC 只登兩個約束，不登 pro 人數點值；ID 退役"),
        ("SRC_OAI", "SRC_OAI_097", "付費企業用戶：2026-02", "711cfec 新增", "V9 掃描 seats：錨點"),
        ("SRC_OAI", "SRC_OAI_098", "研發費用中付 Microsoft 部分：2025（10.59）", "711cfec 新增", "V6：10.59 進 SRC_OAI"),
        ("SRC_OAI", "SRC_OAI_099", "自有園區資本支出例：Project Camellia", "711cfec 新增", "V9 掃描 ownedCapex：錨點"),
        ("SRC_OAI", "SRC_OAI_100", "ChatGPT 每日訊息數：2025-07", "711cfec 新增", "V9 掃描 tasksPerDay：錨點"),
        ("SRC_OAI", "SRC_OAI_101", "Pro 用戶倍數：2026 對 2025（倍增）", "本版新增", "V12 約束一"),
        ("SRC_OAI", "SRC_OAI_102", "Pro 占付費訂閱總數：2026 上限（<1%）", "本版新增", "V12 約束二（只有『高』欄，無點值）"),
        ("Inputs", "INP_057", "FY 平均用戶數：free 2026（950）", "ebe08a4 有；711cfec 移除", "V9：2026 有觀測（WAU）→公式（Derived_V9 V01）"),
        ("Inputs", "INP_072", "FY 平均用戶數：plus 2025（30）", "ebe08a4 有；711cfec 移除", "V9：2025 有觀測→公式（V02）"),
        ("Inputs", "INP_080", "FY 平均用戶數：pro 2025（0.5）", "ebe08a4 有；711cfec 移除", "V9／V12：公式（V04）"),
        ("Inputs", "INP_081", "FY 平均用戶數：pro 2026（1.0）", "ebe08a4 有；711cfec 移除", "V9／V12：公式（V03）"),
        ("Inputs", "INP_195", "非算力營運費用占營收比 2025（1.216）", "ebe08a4 有；711cfec 移除", "V9：2025 由財報數據推得→公式（P4 實作）"),
        ("Inputs", "INP_231", "每 GW 年合約價（r1 新增重複列）", "ebe08a4 有；711cfec 移除", "V8：合約價只留遷入的 leasePricePerGWyr（INP_229）"),
        ("Inputs", "INP_229", "每 GW 年合約價（leasePricePerGWyr）", "711cfec 更名、上限 16→20", "V8"),
        ("Inputs", "INP_183", "合約期間迄：Microsoft Azure", "本版：Analogy→Assumed，值 2030，區間 2029–2032", "V13（同 ID，值不變）"),
        ("Inputs", "INP_184", "合約總額（估計）：Cerebras", "本版：原『總額上緣 25』改為 20（20–25），Assumed", "V13（同 ID，更名）"),
        ("Inputs", "INP_185", "合約期間迄：Cerebras", "本版：Analogy→Assumed，值 2031，區間 2029–2032", "V13（同 ID）"),
        ("Inputs", "INP_234", "換算係數：free FY2026 平均 ÷ 2026-02 WAU", "23a4e04 新增", "E6：V9 公式的換算係數進 Inputs（1.0556，0.844–1.267，Assumed）"),
        ("Inputs", "INP_235", "換算係數：Plus FY2025 平均 ÷ 2025-07 付費訂閱", "23a4e04 新增", "E6（0.857，0.8–0.943，Assumed）"),
        ("Inputs", "INP_236", "Pro 占付費訂閱總數比例（2026）", "23a4e04 新增", "V12（0.8%，0.5–1.0%，Assumed）"),
        # ── 本版（r3 補做＋V15；基準＝23a4e04）──
        ("SRC_OAI", "SRC_OAI_052", "反向目標營收：2026", "23a4e04 有；本版移除", "E8e：與 deckJul2026 同一來源（FT 2026-09-18），合併為 SRC_OAI_043（reverse 引用）；ID 退役"),
        ("SRC_OAI", "SRC_OAI_053", "反向目標營收：2030", "23a4e04 有；本版移除", "E8e：合併為 SRC_OAI_044；ID 退役"),
        ("SRC_OAI", "SRC_OAI_069", "合約容量（另一口徑）：AWS Trainium＝5", "23a4e04 有；本版移除", "E8g：『5』改公式（Derived_V9 V26＝103＋104）；ID 退役"),
        ("SRC_OAI", "SRC_OAI_081", "股權融資：2026 無條件部分＝87", "23a4e04 有；本版移除", "E8j：87 拆為四個 SRC 組成（109–112），加總為公式（V27）；ID 退役"),
        ("SRC_OAI", "SRC_OAI_103", "合約容量（WSJ 口徑）：AWS Trainium 推論 3GW", "本版新增", "E8g：WSJ 2026-07 揭露"),
        ("SRC_OAI", "SRC_OAI_104", "合約容量（WSJ 口徑）：AWS Trainium 訓練（VR）2GW", "本版新增", "E8g：WSJ 2026-07 揭露"),
        ("SRC_OAI", "SRC_OAI_105", "Nvidia 意向：10GW", "本版新增", "E8h：比較用，不計入現金流"),
        ("SRC_OAI", "SRC_OAI_106", "Nvidia 意向：$100B", "本版新增", "E8h：比較用，不計入現金流"),
        ("SRC_OAI", "SRC_OAI_107", "Broadcom 晶片：10GW", "本版新增", "E8h：比較用，不計入現金流"),
        ("SRC_OAI", "SRC_OAI_108", "AMD 晶片：6GW", "本版新增", "E8h：比較用，不計入現金流"),
        ("SRC_OAI", "SRC_OAI_109", "股權融資 2026-03 輪：Amazon 首筆 15", "本版新增", "E8j"),
        ("SRC_OAI", "SRC_OAI_110", "股權融資 2026-03 輪：SoftBank（三期）30", "本版新增", "E8j"),
        ("SRC_OAI", "SRC_OAI_111", "股權融資 2026-03 輪：Nvidia 30", "本版新增", "E8j"),
        ("SRC_OAI", "SRC_OAI_112", "股權融資 2026-03 輪：其他 12", "本版新增", "E8j"),
        ("SRC_OAI", "SRC_OAI_043／044", "管理層營收目標 2026／2030", "同 ID；備註改為 reverse 引用", "E8e"),
        ("SRC_OAI", "SRC_OAI_080", "股權融資：2026-03 輪總額 122", "同 ID；標記改 Interested-party、出處『OpenAI 公告』", "E8i"),
        ("SRC_OAI", "SRC_OAI_022–024", "API 層級對應 top／mid／low", "同 ID；標記改 Verified、出處『OpenAI 價格頁』", "E8b"),
        ("Inputs", "INP_082–085", "FY 平均用戶數：pro 2027–2030（1.5／2.0／2.5／3.0）", "23a4e04 有；本版移除", "V15：改為公式（Derived_V9 V10–V13）；ID 退役"),
        ("Inputs", "INP_086、087", "FY 平均用戶數：pro 低／高情境倍數（0.6／1.2）", "23a4e04 有；本版移除", "V15：區間由比例 0.5–2.0% 傳遞，倍數不再使用；ID 退役"),
        ("Inputs", "INP_237", "API FY2025 平均每分鐘 token（5，3.8–6.1）", "本版新增", "E8a：驅動值進 Inputs（Assumed）"),
        ("Inputs", "INP_238", "每年分鐘數（525,600，定義常數）", "本版新增", "E6／E8a：公式內不含常數"),
        ("Inputs", "INP_239", "單位換算：B→T 除數（1000，定義常數）", "本版新增", "E6／E8a"),
        ("Inputs", "INP_240–243", "Pro 占付費訂閱總數比例 2027–2030（0.8%，0.5–2.0%）", "本版新增", "V15"),
        ("Inputs", "INP_244", "合約容量（採用值）：AWS（Trainium）2（2–5）", "本版新增", "E8g"),
        ("Inputs", "INP_245", "Nvidia $30B 現金比例（1，0–1）", "本版新增", "E8j（沿用 v0.5 S4b）"),
        ("Inputs", "INP_182", "合約期間起：Microsoft Azure", "同 ID：值 2025；低 空→2025；高 空→2026；Analogy 不變；備註改", "V16（Andy 2026-10-04）"),
    ]
    with open(a.out / f"{a.prefix}_row_reconciliation.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["頁", "ID", "項目", "變動", "原因"])
        w.writerows(recon)
    # (b) Decision 標記的 Inputs 列展開
    with open(a.out / f"{a.prefix}_decision_rows.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["ID", "參數", "索引", "值", "原決策", "v0.5 路徑"])
        for r in R.inp_rows:
            if str(r["tag"]).startswith("Decision"):
                w.writerow([r["id"], r["name"], r["index"], r["value"], r["decision"], r["v05"].replace(SEP, ".")])
    # (c) 不遷入葉節點分類小計
    cats = [("說明文字／版本字串／單位字串", ("說明", "版本", "單位字串", "公司名稱", "文字說明")),
            ("結構清單（方案、層級、營收線、選項、充分條件組合）", ("清單", "結構", "營收線", "選項", "充分條件")),
            ("決策記錄（以公式結構體現）", ("決策記錄", "方法說明", "定義", "校準方式")),
            ("null／N／A（未揭露、不適用）", ("null", "N/A")),
            ("v0.4 已廢止或已過時", ("廢止", "過時")),
            ("曆年制設定（工作單已定）", ("曆年制",))]
    cnt = Counter()
    rows_skip = []
    for p, (k, ref, why) in R.dest.items():
        if k != "SKIP":
            continue
        c = next((n for n, keys in cats if any(x in why for x in keys)), "其他")
        cnt[c] += 1
        rows_skip.append((p.replace(SEP, "."), c, why))
    with open(a.out / f"{a.prefix}_skip_summary.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["分類", "葉節點數"])
        w.writerows(cnt.most_common())
        w.writerow([])
        w.writerow(["v0.5 路徑", "分類", "理由"])
        w.writerows(rows_skip)
    # (d) V9／V12 新增 SRC_OAI 列的初評一致性
    v9rows = []
    for i, r in enumerate(R.src_rows):
        if "r2/V9" in r["v05"] or "src文字（V9" in r["v05"] or r["metric"].startswith(("Pro ", "研發費用中付")):
            c = source_rules.classify(r, R.src_rows[i - 1]["source"])
            v9rows.append((r["id"], r["metric"], r["source"][:36], c["grade"], c["stance"], c["why"][:20], c["hand"]))
    with open(a.out / f"{a.prefix}_v9_src_rows.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["ID", "指標", "出處", "等級", "立場", "立場說明", "一手／二手"])
        w.writerows(v9rows)
    print("skip", dict(cnt), "v9 rows", len(v9rows))
    print("\n".join(lines))
    print("tagged objects", len(rows), "v9", len(R.v9))


if __name__ == "__main__":
    main()
