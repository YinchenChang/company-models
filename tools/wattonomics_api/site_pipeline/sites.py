"""Wattonomics 全球 AI 園區候選表（API 版，2026-10-28 起）。

兩段：
1. 種子＋擴充：從 evidence_raw 各 csv 與 B_summary.md 已提到的園區取種子，再以一則一般 API 請求（附 web search）
   擴充到 100–150 個全球 AI 園區（去重）。請求可直接給 sites 清單而略過本段。
2. 逐園區查核（Batch，每園區一則請求，附 web search，max_uses 上限）：MW（IT／設施／未說明）、營運者、使用者、地點、
   公告上線時點、分期、延遲報導、來源 URL 與日期、標記（[Verified]／[Interested-party]／[Derived]／[Assumed]）。
輸出候選表 out/<名>_candidates.csv＋.md 與 out/<名>_sites.csv（園區清單），**不寫進 Excel**。

請求檔 tools/wattonomics_api/site_pipeline/requests/<名>.json：
  {"max_sites": 150, "expand": true, "target_sites": 130, "sites": [], "searches_per_site": 5,
   "expand_searches": 25, "cost_cap_usd": 25, "evidence_dir": null}
  冒煙測試：{"max_sites": 3, "expand": false, "searches_per_site": 3, "cost_cap_usd": 1}
"""
import csv
import glob
import json
import os
import re
import shutil
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import common as C  # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
TOOL = "site_pipeline"
DEFAULT_CAP = 25.0
MAX_SEARCH_PER_SITE = 8          # 硬上限（請求檔再大也截到此值）
MAX_EXPAND_SEARCH = 40
MAX_SITES = 150
EST_PER_SITE_USD = {"tok": 0.12}   # 事前估計：每園區 token 約 US$0.12（Batch；搜尋結果計入輸入），另加搜尋次數 × 0.01
NON_SITE = re.compile(r"全公司|全體|全部|總量|總容量|資本支出|法說|晶片|承諾|合約總|目標|估|平均|在役|租用|簽約|預期|FY\d|20\d\d ?Q|曆年")


def seeds(ev_dir):
    """從 csv 的園區欄與 B_summary 表格取種子名稱（程式粗篩，模型在擴充段去重）。"""
    out = []
    for p in sorted(glob.glob(f"{C.ROOT}/{ev_dir}/*.csv")):
        with open(p, encoding="utf-8-sig", newline="") as f:
            for r in csv.DictReader(f):
                name = (r.get("site_or_scope") or r.get("site_or_contract") or "").strip()
                if not name or NON_SITE.search(name):
                    continue
                out.append({"name": name, "player": (r.get("player") or "").strip(),
                            "operator": (r.get("operator") or "").strip(), "user": (r.get("user") or "").strip(),
                            "from": os.path.basename(p)})
    summ = f"{C.ROOT}/{ev_dir}/B_summary.md"
    if os.path.exists(summ):
        for line in open(summ, encoding="utf-8"):
            if line.startswith("|") and not line.startswith("|---"):
                cell = line.strip("|\n").split("|")[0].strip()
                if cell and cell not in ("項目", "營運者", "園區") and not NON_SITE.search(cell):
                    out.append({"name": cell, "player": "", "operator": "", "user": "", "from": "B_summary.md"})
    dedup = {}
    for s in out:
        dedup.setdefault(s["name"], s)
    return list(dedup.values())


SYS_EXPAND = """你在建立「全球 AI 資料中心園區」清單（Wattonomics 全市場 GW 帳）。給你從研究證據檔取出的種子名稱（可能重複、混有合約名或非園區項目）。
任務：
1. 種子去重、合併同一園區的不同寫法；剔除不是實體園區的項目（公司總量、合約、晶片採購）。
2. 用 web search 補足到目標數量：涵蓋美國（各州）、中東、歐洲、亞洲（含中國、印度、日韓、東南亞）、拉美、澳洲；
   以 AI／GPU 訓練或推論為主、規劃或在建或已上線、≥100MW（或公告將擴到 ≥100MW）的園區優先。
3. 每個園區一行 JSON（JSON Lines，不要其他文字）：
{"name":園區名（英文常用名＋城市）,"location":"城市, 州/省, 國家","player":主要相關公司,"operator":營運者（若知）,"user":主要使用者（若知）,"seed":true/false（是否來自種子）}"""

SYS_SITE = """你是資料中心園區研究員。用 web search 查一個 AI 資料中心園區的最新公開資訊（以最近 12 個月為主），只根據搜尋到的來源回答，不得用記憶補值。
輸出 JSON Lines，不要其他文字。每個「分期」（或全園區）一行：
{"site":園區名,"phase":分期名（全園區寫"全園區"）,"mw":數字或null,"mw_basis":"IT"/"設施"/"未說明","mw_status":"上線"/"在建"/"規劃"/"簽約",
 "operator":營運者,"user":使用者（多個用／分隔；未揭露寫"未揭露"）,"location":"城市, 州/省, 國家",
 "announced_online":公告上線時點（YYYY-MM 或 YYYY Qn 或 YYYY）,"delay":延遲報導（無則空字串；有則寫延多久與原因）,
 "source_url":來源網址,"source_date":來源日期 YYYY-MM-DD,"tag":"[Verified]"/"[Interested-party]"/"[Derived]"/"[Assumed]",
 "who_said":說話者（公司本人／記者推論／第三方機構）,"note":補充（繁體中文 60 字內）}
標記規則：[Verified]＝監管文件、財報、政府許可或獨立機構數據；[Interested-party]＝營運者、使用者、開發商或其投資人的說法（含新聞稿）；
[Derived]＝你由來源數字推算（在 note 寫算式）；[Assumed]＝來源沒有、你推定的（盡量少用）。
MW 沒說 IT 還是設施，一律寫「未說明」，不要自行換算。同一數字有多個來源時取最權威、最新的一個，另一個寫在 note。
查不到任何資訊時，輸出一行 mw 為 null、note 寫「查無公開資訊」。"""


def expand(seed_list, req, usage, cap):
    n = min(int(req.get("expand_searches", 25)), MAX_EXPAND_SEARCH)
    target = int(req.get("target_sites", 130))
    seed_txt = "\n".join(f"- {s['name']}｜{s['player']}｜營運 {s['operator']}｜使用 {s['user']}" for s in seed_list)
    params = {"model": C.MODEL, "max_tokens": 32000, "system": SYS_EXPAND,
              "tools": [C.web_search_tool(n)],
              "messages": [{"role": "user", "content": f"目標數量：{target} 個（100–150 之間）。\n\n## 種子（{len(seed_list)} 筆）\n{seed_txt}"}]}
    if cap < 2:
        raise RuntimeError("預算不足以執行擴充段（需約 US$1–2）")
    m = C.standard_message(params, usage)
    sites = C.parse_jsonl(C.text_of(m))
    dedup = {}
    for s in sites:
        k = re.sub(r"\W+", "", str(s.get("name", "")).lower())
        if k:
            dedup.setdefault(k, s)
    return list(dedup.values())


def write_sites(name, sites):
    os.makedirs(f"{HERE}/out", exist_ok=True)
    with open(f"{HERE}/out/{name}_sites.csv", "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["name", "location", "player", "operator", "user", "seed"])
        for s in sites:
            w.writerow([s.get(k, "") for k in ("name", "location", "player", "operator", "user", "seed")])


COLS = ["site", "phase", "mw", "mw_basis", "mw_status", "operator", "user", "location", "announced_online", "delay",
        "source_url", "source_date", "tag", "who_said", "note"]


def write_candidates(name, sites, res, usage, skipped):
    rows = []
    for i, s in enumerate(sites):
        got = C.parse_jsonl(res.get(f"s{i}", ""))
        if not got:
            got = [{"site": s.get("name"), "note": "未取得結果" if f"s{i}" not in res else "輸出無法解析"}]
        for g in got:
            g.setdefault("site", s.get("name"))
            g["_query"] = s.get("name")
            rows.append(g)
    with open(f"{HERE}/out/{name}_candidates.csv", "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["query_site"] + COLS)
        for r in rows:
            w.writerow([r["_query"]] + [r.get(c, "") if r.get(c) is not None else "" for c in COLS])
    tags, basis, delayed, mw_known = {}, {}, 0, 0
    for r in rows:
        tags[r.get("tag") or "無"] = tags.get(r.get("tag") or "無", 0) + 1
        basis[r.get("mw_basis") or "無"] = basis.get(r.get("mw_basis") or "無", 0) + 1
        delayed += bool(str(r.get("delay") or "").strip())
        mw_known += isinstance(r.get("mw"), (int, float))
    md = [f"# Wattonomics AI 園區候選表｜{name}", "",
          f"- 園區 {len(sites)} 個 → 候選列 {len(rows)}（有 MW 數字 {mw_known}；有延遲報導 {delayed}）",
          "- MW 口徑：" + "、".join(f"{k} {v}" for k, v in basis.items()),
          "- 標記：" + "、".join(f"{k} {v}" for k, v in tags.items()),
          f"- 用量：input {usage.d['input_tokens']:,}／output {usage.d['output_tokens']:,} tokens；web search {usage.d['web_search_requests']} 次；估計 US${usage.usd:.2f}",
          ]
    if skipped:
        md.append(f"- **超出本次費用上限未查：{skipped} 個園區**")
    md += ["", "| 園區 | 分期 | MW | 口徑 | 狀態 | 營運者 | 使用者 | 上線 | 延遲 | 標記 |", "|---|---|---|---|---|---|---|---|---|---|"]
    for r in rows[:400]:
        md.append("| " + " | ".join(str(r.get(c) or "")[:40].replace("|", "／") for c in
                                     ("site", "phase", "mw", "mw_basis", "mw_status", "operator", "user", "announced_online", "delay", "tag")) + " |")
    md += ["", "候選值不是模型數值；進 Wattonomics Excel 前須逐列對來源核對（不直接寫入 Excel）。來源網址見同名 csv。"]
    open(f"{HERE}/out/{name}_candidates.md", "w", encoding="utf-8").write("\n".join(md) + "\n")
    return len(rows), "\n".join(md[:6])


def main():
    if not C.API_KEY:
        print("未設定 ANTHROPIC_API_KEY，結束。")
        return
    for path in sorted(glob.glob(f"{HERE}/requests/*.json")):
        name = os.path.splitext(os.path.basename(path))[0]
        req = json.load(open(path, encoding="utf-8"))
        stamp = time.strftime("%Y%m%dT%H%M%SZ", time.gmtime())
        usage = C.Usage(TOOL, name)
        try:
            cap, _ = C.run_cap(float(req.get("cost_cap_usd", DEFAULT_CAP)))
            if cap <= 0.2:
                raise RuntimeError(f"預算已用盡（本次可用 US${cap:.2f}），未執行")
            ev = C.evidence_dir(req.get("evidence_dir"))
            max_sites = min(int(req.get("max_sites", MAX_SITES)), MAX_SITES)
            if req.get("_sites"):                                   # 續接：沿用上次的園區清單
                sites = req["_sites"]
            elif req.get("sites"):
                sites = [s if isinstance(s, dict) else {"name": s} for s in req["sites"]]
            else:
                sd = seeds(ev)
                print("種子", len(sd))
                sites = expand(sd, req, usage, cap) if req.get("expand", True) else \
                    [{"name": s["name"], "player": s["player"], "operator": s["operator"], "user": s["user"], "seed": True} for s in sd]
            sites = sites[:max_sites]
            write_sites(name, sites)
            nsearch = min(int(req.get("searches_per_site", 5)), MAX_SEARCH_PER_SITE)
            per = EST_PER_SITE_USD["tok"] + nsearch * C.SEARCH_USD
            room = cap - usage.usd
            n_ok = max(0, min(len(sites), int(room / per)))
            skipped = len(sites) - n_ok
            print(f"園區 {len(sites)}；每園區搜尋上限 {nsearch}；事前估計每園區 US${per:.2f}；可查 {n_ok}")
            params = {}
            for i, s in enumerate(sites[:n_ok]):
                hint = "｜".join(f"{k}：{s[k]}" for k in ("location", "player", "operator", "user") if s.get(k))
                params[f"s{i}"] = {"model": C.MODEL, "max_tokens": 6000, "system": SYS_SITE,
                                   "tools": [C.web_search_tool(nsearch)],
                                   "messages": [{"role": "user", "content": f"園區：{s.get('name')}\n已知線索：{hint or '無'}\n今天：{time.strftime('%Y-%m-%d')}"}]}
            res, state = C.run_requests(params, usage, req.get("_state"))
            if state.get("status") == "pending":
                req["_state"], req["_sites"] = state, sites
                json.dump(req, open(path, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
                open(f"{HERE}/out/{name}_candidates.md", "w", encoding="utf-8").write(
                    f"# 園區查核處理中｜{name}\n\n批次 {state['batch_id']} 尚未完成；重新觸發會續接，不重送。\n")
                usage.write(f"{HERE}/out/{name}_usage_{stamp}.json", {"status": "pending"})
                continue
            os.makedirs(f"{HERE}/out/raw", exist_ok=True)
            json.dump(res, open(f"{HERE}/out/raw/{name}.json", "w", encoding="utf-8"), ensure_ascii=False)
            nrows, summary = write_candidates(name, sites, res, usage, skipped)
            usage.write(f"{HERE}/out/{name}_usage_{stamp}.json", {"status": "done", "sites": len(sites),
                                                                  "sites_queried": n_ok, "rows": nrows})
            print(name, "候選列", nrows, f"US${usage.usd:.2f}")
            C.pr_comment(summary + f"\n\n結果：`tools/wattonomics_api/site_pipeline/out/{name}_candidates.csv`（參考用，不阻擋合併）")
            os.makedirs(f"{HERE}/requests/done", exist_ok=True)
            shutil.move(path, f"{HERE}/requests/done/{name}.json")
        except Exception as e:
            os.makedirs(f"{HERE}/out", exist_ok=True)
            open(f"{HERE}/out/{name}_candidates.md", "w", encoding="utf-8").write(
                f"# 園區查核失敗｜{name}\n\n{type(e).__name__}: {e}\n\n請求檔保留在 requests/，修正後重推即重跑。\n")
            usage.write(f"{HERE}/out/{name}_usage_{stamp}.json", {"status": "failed", "error": str(e)[:300]})
            print("失敗", name, e)


if __name__ == "__main__":
    main()
