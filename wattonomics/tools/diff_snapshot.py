# Wattonomics 快照逐名比對：既有 WIF_ 名稱的值是否變動（新版只加頁時，差異清單必須為空）。
# 用法：python3 wattonomics/tools/diff_snapshot.py <舊 snapshot.json> <新 snapshot.json 或重算後 xlsx> [--json <輸出>] [--allow 前綴,...]
#   第二個參數給 xlsx 時直接讀活頁簿（export_snapshot.read）。WGOV_ 治理計數另列，不算 WIF_ 差異。
#   --allow：允許變動的名稱前綴（例如合併新證據後的 WIF_Gap_）；其餘任何 WIF_ 差異以代碼 1 結束。
#   浮點雜訊 < 1E-9 另列，不算差異（工作單 §7）。
#   v0.2：舊有新無的名稱另列「刪除」（--allow-removed 時不算差異）；--top N 列出既有名稱變動前 N 大（絕對差）與前綴分類。
import json, os, sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)


def load(path):
    if path.endswith(".xlsx"):
        from export_snapshot import read
        return read(path)[0]
    return json.load(open(path, encoding="utf-8"))["names"]


def main():
    args = sys.argv[1:]
    out = allow = None
    if "--json" in args:
        i = args.index("--json"); out = args[i + 1]; del args[i:i + 2]
    if "--allow" in args:
        i = args.index("--allow"); allow = tuple(args[i + 1].split(",")); del args[i:i + 2]
    allow_removed = "--allow-removed" in args
    if allow_removed: args.remove("--allow-removed")
    top = 0
    if "--top" in args:
        i = args.index("--top"); top = int(args[i + 1]); del args[i:i + 2]
    old, new = load(args[0]), load(args[1])
    diff, noise, gov, allowed = [], [], [], []
    removed = sorted(k for k in old if k not in new)
    for k, v in old.items():
        if k not in new:
            if not allow_removed and k.startswith("WIF_"): diff.append((k, v, "<缺>"))
            continue
        n = new.get(k, "<缺>")
        if n == v: continue
        if isinstance(v, (int, float)) and isinstance(n, (int, float)) and abs(n - v) < 1e-9:
            noise.append((k, v, n)); continue
        if not k.startswith("WIF_"): gov.append((k, v, n)); continue
        (allowed if allow and k.startswith(allow) else diff).append((k, v, n))
    added = sorted(k for k in new if k not in old)
    print(f"舊 {len(old)} 名、新 {len(new)} 名；刪除 {len(removed)}；既有 WIF_ 差異 {len(diff)}；允許前綴內變動 {len(allowed)}；浮點雜訊 {len(noise)}；"
          f"治理計數變動 {len(gov)}；新增名稱 {len(added)}（WIF_Gap_ {sum(1 for k in added if k.startswith('WIF_Gap_'))}）")
    if top:
        import re
        from collections import Counter
        ch = [d for d in diff + allowed if d[2] != "<缺>"]
        cls = lambda k: re.sub(r"_(\d{4})$", "", re.sub(r"_(msft|googl|amzn|meta|orcl|crwv|nbis|wyfi|xai|oai|ant|neo|sov|cn|oth)(?=_|$)", "_<p>", k)).split("_<p>")[0]
        print("  刪除前綴：", dict(Counter(re.sub(r"_\d{4}$", "", k).split("_A_")[0].split("_B_")[0].split("_C_")[0] for k in removed).most_common(15)))
        print("  變動分類（前 20）：", dict(Counter(cls(k) for k, _, _ in ch).most_common(20)))
        num = [(k, v, n, abs(n - v)) for k, v, n in ch if isinstance(v, (int, float)) and isinstance(n, (int, float))]
        for k, v, n, d_ in sorted(num, key=lambda x: -x[3])[:top]: print(f"  變動 {k}: {v:.4g} → {n:.4g}（{n - v:+.4g}）")
        for k, v, n in [x for x in ch if not (isinstance(x[1], (int, float)) and isinstance(x[2], (int, float)))][:15]: print(f"  文字變動 {k}: {v} → {n}")
    for d in diff[:50]: print("  差異", d)
    for d in gov: print("  治理", d)
    if out:
        json.dump(dict(removed=removed, diff=diff, allowed=allowed, noise=noise, governance=gov, added=added), open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=0)
    sys.exit(1 if diff else 0)


if __name__ == "__main__":
    main()
