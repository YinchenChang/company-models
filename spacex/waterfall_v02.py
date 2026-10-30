# v0.1 → v0.2 目標價歸因瀑布：以 SPCX_STAGE＝0…6 逐項套用改動（0＝v0.1 公式與輸入、在 v0.2 事實檔上重建），每一階段整本 LibreOffice 重算，
# 記錄加權目標價、評等、收支關卡 B、FY2028 S2 比值 B、讀法 1 目標價；相鄰兩階段之差即該項貢獻（合計＝v0.2 − v0.1）。
# 第 0 階段另以 SPCX_NO_G2CAL 不影響（G2 校準只在 Sens）。
# 用法：python3 spacex/waterfall_v02.py <工作資料夾> <輸出 json>（JOBS＝平行數，預設 2）
#   WF=v021（預設）：v0.2 → v0.2.1 兩步（SPCX_STAGE 7 → 8 → 9）；WF=v02：v0.1 → v0.2（0…7，data/waterfall_v02_20261030.json）
import json, os, subprocess, sys, shutil, tempfile
from concurrent.futures import ThreadPoolExecutor
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
WD, OUT = sys.argv[1], sys.argv[2]
os.makedirs(WD, exist_ok=True)
STAGES_V02 = [(0, "v0.1（重建；應等於 v0.1 成品 28.76）"),
          (1, "小修 (a)：分部加總讀法 2 刪除無對應扣除的自用 AI 營運成本加回（讀法 1 改扣）"),
          (2, "小修 (b)：FY2027 年底淨現金只扣融資租賃年底餘額（不重扣已付本金）"),
          (3, "G5：424B4 原文（2024 年底名目算力 0.15 → 0.3、2025 年底 0.8 不變；其餘 FY2025 原文只影響實績欄）"),
          (4, "G1：Q4 26 合併資本支出 37.5（SI＝37.5 − Space − Connectivity；取代 Q3 季速）"),
          (5, "G1：2026 年底名目算力 2.75 GW（Q4 新增 1.05；取代 pace × 0.25 推得的 2.11）"),
          (6, "G3（一修）：FY2027–28 pace＝(37.5 × 9M26 SI 占比 83.2%＋以融資取得 PP&E 1.307) × 4 ÷ 每 GW 投資＝2.98"),
          (7, "G3 一致性（二修）：pace 改用與 Q4 相同拆法＝(37.5 − Space − Connectivity 3.437＋1.307) × 4 ÷ 43.64＝3.24")]
STAGES_V021 = [(7, "v0.2（重建；應等於 v0.2 成品 34.83）"),
               (8, "v0.2.1 第 1 點：pace 的 SI 投資季速改與 Q4 同法＝Q4 SI 現金資本支出 ÷ (1 − fin_sh)（取代 ＋以融資取得 PP&E 9M 季均 1.307）"),
               (9, "v0.2.1 第 2 點：pace 分母改 FY2027 新增世代組合每 GW 投資（am_27：GB300／VR200 各半；取代 GB300 43.64）")]
STAGES = STAGES_V021 if os.environ.get("WF", "v021") == "v021" else STAGES_V02


def run(stage):
    raw = os.path.join(WD, f"s{stage}.xlsx")
    env = dict(os.environ); env["SPCX_STAGE"] = str(stage); env["ROWMAP"] = raw + ".rowmap.json"
    subprocess.run([sys.executable, os.path.join(HERE, "build_spacex.py"), raw], check=True, stdout=subprocess.DEVNULL, env=env)
    out = os.path.join(WD, f"out{stage}"); os.makedirs(out, exist_ok=True)
    prof = tempfile.mkdtemp(prefix="lo_wf_", dir=WD)
    subprocess.run(["soffice", f"-env:UserInstallation=file://{prof}", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", out, raw],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, env={**os.environ, "LC_ALL": "C.UTF-8"}, timeout=900)
    shutil.rmtree(prof, ignore_errors=True)
    rm = json.load(open(raw + ".rowmap.json", encoding="utf-8"))
    wb = openpyxl.load_workbook(os.path.join(out, os.path.basename(raw)), data_only=True)
    V, CF, SG, TH = wb["Val"], wb["CF"], wb["Seg"], wb["Thesis"]
    vr = rm["VR"]
    return {"tp": V.cell(vr["tp"], 2).value, "rec": V.cell(vr["rec"], 2).value, "tp_r1": V.cell(vr["tp_r1"], 2).value,
            "tpdcf": V.cell(vr["tpdcf"], 2).value, "tpcomp": V.cell(vr["tpcomp"], 2).value,
            "gateB": CF.cell(rm["GATE"], 3).value, "s2B": SG[f"I{rm['SR']['B']['s2']}"].value,
            "gw26": TH[f"F{rm['T']['end']}"].value, "gw27": TH[f"H{rm['T']['end']}"].value, "gw28": TH[f"I{rm['T']['end']}"].value,
            "inv27": TH[f"H{rm['T']['si_inv']}"].value, "q4inv": TH[f"F{rm['T']['q4inv']}"].value,
            "pace": wb["In"].cell(rm["IN_ROW"]["pace"], 5).value, "si_cx_rr": wb["In"].cell(rm["IN_ROW"]["si_cx_rr"], 5).value, "plugB": CF.cell(rm["GATE"] + 2, 3).value}


with ThreadPoolExecutor(max_workers=int(os.environ.get("JOBS", "2"))) as ex:
    res = dict(zip([s for s, _ in STAGES], ex.map(run, [s for s, _ in STAGES])))
rows = []
for i, (s, lab) in enumerate(STAGES):
    o = res[s]
    d = None if i == 0 else o["tp"] - res[STAGES[i - 1][0]]["tp"]
    rows.append({"stage": s, "label": lab, **o, "delta_tp": d})
    print(s, round(o["tp"], 2), "" if d is None else f"{d:+.2f}", o["rec"], round(o["gateB"], 1), round(o["s2B"], 3), lab, flush=True)
tot = res[STAGES[-1][0]]["tp"] - res[STAGES[0][0]]["tp"]
assert abs(sum(r["delta_tp"] for r in rows[1:]) - tot) < 1e-9
json.dump({"rows": rows, "total_delta": tot}, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
