# W3（2026-10-08）：v4.5 → v4.6 前後對照的取數（給 scripts/build_compare.py）。
# 三個 Excel × 三情境，各以 xlx.py（LibreOffice 重算；情境以「情境選擇」切換）讀成「工作表|列名稱 → 5 期值」：
#   v45＝v4.5 成品（dist/，git 歷史取出）；v45pm＝舊方法組合的副本建置（scripts/verify_legacy.sh 產生，與 v4.5 成品 0 差異，只為取「每MW經濟性」頁）；v46＝v4.6 成品。
# 檢查：v45pm 與 v45 共有的每一個數值格（全部工作表、三情境）相同（相對誤差 ≤ 1e-9），否則失敗——v4.5 欄的每 MW 數字因此可追溯到 v4.5 成品。
# 用法：python3 scripts/compare_gather.py --v45 v4.5.xlsx --v45pm 舊方法副本.xlsx --v46 v4.6.xlsx --out 取數.json
import sys, os, json, argparse, subprocess
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ap = argparse.ArgumentParser()
for k in ('v45', 'v45pm', 'v46', 'out'): ap.add_argument('--' + k, required=True)
a = ap.parse_args()
res, fails = {}, []
for tag in ('v45', 'v45pm', 'v46'):
    res[tag] = {}
    for s in (1, 2, 3):
        o = os.path.join(ROOT, 'out', 'w3', f'xl_{tag}_{s}.json')
        r = subprocess.run(['python3', os.path.join(ROOT, 'xlx.py'), os.path.abspath(getattr(a, tag)), str(s), o], capture_output=True, text=True)
        if r.returncode: sys.exit(f'xlx 失敗 {tag} {s}：{r.stdout}{r.stderr}')
        res[tag][s] = json.load(open(o, encoding='utf-8'))
n = 0
for s in (1, 2, 3):
    A, B = res['v45'][s], res['v45pm'][s]
    for k in A:
        if k.startswith('檢查_版本紀錄|') or k not in B: continue
        for x, y in zip(A[k], B[k]):
            if isinstance(x, (int, float)) and not isinstance(x, bool):
                n += 1
                if not isinstance(y, (int, float)) or abs(x - y) > 1e-9 * max(1, abs(x)): fails.append(f'情境 {s} {k}：{x} ≠ {y}')
json.dump({k: {str(s): v for s, v in d.items()} for k, d in res.items()}, open(a.out, 'w', encoding='utf-8'), ensure_ascii=False)
print(f'取數：3 檔 × 3 情境；v4.5 成品 vs 舊方法副本 共有數值格 {n} 個，不一致 {len(fails)}')
for f in fails[:20]: print('  ', f)
sys.exit(1 if fails or n == 0 else 0)
