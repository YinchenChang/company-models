"""WhiteFiber 目標價變動拆解 v0.4.0（全公司 × 新雲同業中位數，5.39）→ v0.4.1（Comparable 分部加總；Andy 2026-11-15）→ data/decomp_v50.json
  ① 同業法改分部加總（託管分部 × 上市資料中心同業、雲端分部 × 新雲同業）：sotp_on 0 → 1；其餘輸入不變，所以 ① 即總變動。
v0.2.1 → v0.4.0 的拆解（−10.48：① 新樣板＋Q3 實績＋Tokenomics v22.0 −10.63；② 現價 7.80 → 9.21 +0.14）見 README v0.4.0 節。"""
import datetime, json, os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import runner_v50 as RN

VOLD = 5.3855368551  # v0.4.0 Val 加權目標價（dist/_待刪除/20261115_WhiteFiber個股模型_v0.4.0.xlsx）
res = RN.run([("s0", {"sotp_on": 0}, {}), ("s1", {}, {})], 2)
tp0, tp1 = res["s0"]["tp"], res["s1"]["tp"]
steps = [{"id": "sotp", "label": "① 同業法改分部加總（託管 × 資料中心同業、雲端 × 新雲同業）", "delta": tp1 - tp0, "tp": tp1}]
out = {"date": datetime.date.today().isoformat(), "vold": VOLD, "vnew": tp1, "total": tp1 - VOLD, "steps": steps,
       "check": sum(x["delta"] for x in steps) - (tp1 - VOLD), "reproduce_v040": {"tp": tp0, "rec": res["s0"]["rec"], "diff": tp0 - VOLD}}
json.dump(out, open(os.path.join(RN.HERE, "data", "decomp_v50.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print(json.dumps(out, ensure_ascii=False, indent=1))
