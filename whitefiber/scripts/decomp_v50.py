"""WhiteFiber 目標價變動拆解 v0.2.1（舊 JS 引擎，15.87）→ v0.4.0（CoreWeave 新雲樣板重建）。兩段相加＝總變動 → data/decomp_v50.json
  ① 方法與資料：新樣板＋Q3 實績＋Tokenomics v22.0（以舊現價 7.80 計）　② 現價 7.80 → 9.21（2026-11-10）
舊引擎與新樣板結構不同，① 不再細拆，README 列主要差異。"""
import datetime, json, os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import runner_v50 as RN

VOLD, PX0 = 15.87, 7.80
res = RN.run([("s1", {"px": PX0}, {}), ("s2", {}, {})], 2)
tp1, tp2 = res["s1"]["tp"], res["s2"]["tp"]
steps = [{"id": "m", "label": "① 新樣板＋Q3 實績＋Tokenomics v22.0（以舊現價 7.80 計）", "delta": tp1 - VOLD, "tp": tp1},
         {"id": "px", "label": "② 現價 7.80 → 9.21（2026-11-10）", "delta": tp2 - tp1, "tp": tp2}]
out = {"date": datetime.date.today().isoformat(), "vold": VOLD, "vnew": tp2, "total": tp2 - VOLD, "steps": steps,
       "check": sum(x["delta"] for x in steps) - (tp2 - VOLD), "full_at_old_px": {"tp": tp1, "rec": res["s1"]["rec"], "up": tp1 / PX0 - 1}}
json.dump(out, open(os.path.join(RN.HERE, "data", "decomp_v50.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print(json.dumps(out, ensure_ascii=False, indent=1))
