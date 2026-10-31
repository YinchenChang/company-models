# Wattonomics 快照逐名比對：既有 WIF_ 名稱的值是否變動（新版只加頁時，差異清單必須為空）。
# 用法：python3 wattonomics/tools/diff_snapshot.py <舊 snapshot.json> <新 snapshot.json 或重算後 xlsx> [--json <輸出>] [--allow 前綴,...]
#   第二個參數給 xlsx 時直接讀活頁簿（export_snapshot.read）。WGOV_ 治理計數另列，不算 WIF_ 差異。
#   --allow：允許變動的名稱前綴（例如合併新證據後的 WIF_Gap_）；其餘任何 WIF_ 差異以代碼 1 結束。
#   浮點雜訊 < 1E-9 另列，不算差異（工作單 §7）。
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
    old, new = load(args[0]), load(args[1])
    diff, noise, gov, allowed = [], [], [], []
    for k, v in old.items():
        n = new.get(k, "<缺>")
        if n == v: continue
        if isinstance(v, (int, float)) and isinstance(n, (int, float)) and abs(n - v) < 1e-9:
            noise.append((k, v, n)); continue
        if not k.startswith("WIF_"): gov.append((k, v, n)); continue
        (allowed if allow and k.startswith(allow) else diff).append((k, v, n))
    added = sorted(k for k in new if k not in old)
    print(f"舊 {len(old)} 名、新 {len(new)} 名；既有 WIF_ 差異 {len(diff)}；允許前綴內變動 {len(allowed)}；浮點雜訊 {len(noise)}；"
          f"治理計數變動 {len(gov)}；新增名稱 {len(added)}（WIF_Gap_ {sum(1 for k in added if k.startswith('WIF_Gap_'))}）")
    for d in diff[:50]: print("  差異", d)
    for d in gov: print("  治理", d)
    if out:
        json.dump(dict(diff=diff, allowed=allowed, noise=noise, governance=gov, added=added), open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=0)
    sys.exit(1 if diff else 0)


if __name__ == "__main__":
    main()
