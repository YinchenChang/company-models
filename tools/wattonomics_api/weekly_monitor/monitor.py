"""Wattonomics 每週變動監測（API 版，2026-10-28 起）。

每週一 09:00（台北）搜尋各玩家上一週（上週一至週日，台北時間）的：新園區、延期、取消、發債、合資、股權募資消息，
輸出「本週變動清單」out/<週一日期>.md（＋ .csv、usage.json）。workflow 把結果 commit 到固定分支 wattonomics-monitor（不寫 main）。
STOP_AFTER（預設 2026-11-10，API credit 到期日）之後自動停止，不呼叫 API。

每玩家一則 Batch 請求，附 web search（max_uses 上限）。
手動或請求檔觸發時可帶參數（請求檔 tools/wattonomics_api/weekly_monitor/requests/<名>.json）：
  {"players": ["Microsoft"], "week_start": "2026-10-19", "searches_per_player": 6, "cost_cap_usd": 4}
"""
import csv
import datetime as dt
import glob
import json
import os
import shutil
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import common as C  # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
TOOL = "weekly_monitor"
STOP_AFTER = os.environ.get("STOP_AFTER", "2026-11-10")
DEFAULT_CAP = 4.0
MAX_SEARCH_PER_PLAYER = 8
PLAYERS = {
    "Microsoft": "Microsoft／Azure（含 Fairwater）",
    "Alphabet": "Alphabet／Google Cloud（含 TPU 園區）",
    "Amazon": "Amazon／AWS（含 Project Rainier）",
    "Meta": "Meta（含 Hyperion、Prometheus）",
    "Oracle": "Oracle／OCI（含 Stargate 站點）",
    "CoreWeave": "CoreWeave",
    "Nebius": "Nebius",
    "OpenAI": "OpenAI（含 Stargate、SB Energy、海外 Stargate）",
    "Anthropic": "Anthropic（含 Fluidstack、Hut 8、TeraWulf 等合作園區）",
    "xAI": "xAI／SpaceX（含 Colossus）",
    "中國大廠": "中國大廠：阿里巴巴、騰訊、字節跳動、百度、華為（及其 AI 資料中心）",
}
CATS = ["新園區", "延期", "取消", "發債", "合資", "股權募資", "其他"]

SYS = """你是 AI 基礎建設新聞監測員。用 web search 找指定公司在指定一週內（含首尾日）首次公開的消息，只收以下六類：
新園區（新公告的資料中心園區或擴建、MW／GW 規模）、延期、取消（含暫停、退出）、發債（公司債、專案融資、資產擔保、私募信用）、
合資（JV、共同開發、算力租約中的合資結構）、股權募資（增資、IPO、可轉債以股權計）。不在該週發布的舊聞不要列。
只根據搜尋到的來源，不得用記憶補值；同一事件多個來源取最權威的一個（監管文件 > 公司新聞稿 > 主流媒體 > 整理站）。
輸出 JSON Lines，每則一行，不要其他文字：
{"player":公司,"category":六類之一,"date":消息日期 YYYY-MM-DD,"headline":一句話摘要（繁體中文 40 字內）,"site":相關園區（無則空字串）,
 "mw":數字或null,"mw_basis":"IT"/"設施"/"未說明","amount_usd_bn":金額（十億美元）或null,"counterparty":對手方,
 "tag":"[Verified]"/"[Interested-party]"/"[Derived]"/"[Assumed]","source_url":網址,"source_date":來源日期,"who_said":公司本人／記者推論／第三方}
該週沒有任何符合的消息時，只輸出一行：{"player":公司,"category":"無","headline":"本週無符合消息"}"""


def week_range(start=None):
    tpe = dt.timezone(dt.timedelta(hours=8))
    today = dt.datetime.now(tpe).date()
    if start:
        s = dt.date.fromisoformat(start)
    else:
        this_mon = today - dt.timedelta(days=today.weekday())
        s = this_mon - dt.timedelta(days=7)
    return today, s, s + dt.timedelta(days=6)


def run(req, name, usage):
    today, ws, we = week_range(req.get("week_start"))
    cap, _ = C.run_cap(float(req.get("cost_cap_usd", DEFAULT_CAP)))
    if cap <= 0.2:
        raise RuntimeError(f"預算已用盡（本次可用 US${cap:.2f}），未執行")
    players = req.get("players") or list(PLAYERS)
    n = min(int(req.get("searches_per_player", 6)), MAX_SEARCH_PER_PLAYER)
    per = 0.12 + n * C.SEARCH_USD
    players = players[: max(0, int(cap / per))]
    params = {f"p{i}": {"model": C.MODEL, "max_tokens": 6000, "system": SYS, "tools": [C.web_search_tool(n)],
                        "messages": [{"role": "user", "content": f"公司：{PLAYERS.get(p, p)}\n期間：{ws} 至 {we}（台北時間）\n今天：{today}"}]}
              for i, p in enumerate(players)}
    print(f"週 {ws}–{we}；玩家 {len(players)}；每玩家搜尋上限 {n}；本次上限 US${cap:.2f}")
    # 每週排程需當天完成：輪詢最多 4.5 小時，仍未完成就取消改一般 API（逾 6 小時規則的保守版）
    res, _ = C.run_requests(params, usage, poll_max_min=10_000, stale_h=min(C.STALE_BATCH_H, 4.5))
    items = []
    for i, p in enumerate(players):
        got = C.parse_jsonl(res.get(f"p{i}", ""))
        if not got:
            got = [{"player": p, "category": "未取得", "headline": "輸出無法解析或請求失敗"}]
        for g in got:
            g["player"] = g.get("player") or p
        items += got
    out = f"{HERE}/out"
    os.makedirs(out, exist_ok=True)
    base = f"{out}/{name}"
    cols = ["player", "category", "date", "headline", "site", "mw", "mw_basis", "amount_usd_bn", "counterparty",
            "tag", "source_url", "source_date", "who_said"]
    with open(base + ".csv", "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(cols)
        for it in items:
            w.writerow(["" if it.get(c) is None else it.get(c) for c in cols])
    real = [it for it in items if it.get("category") in CATS]
    md = [f"# Wattonomics 本週變動清單｜{ws} 至 {we}", "",
          f"- 產生：{today}（台北）；玩家 {len(players)}；符合消息 {len(real)} 則",
          f"- 用量：input {usage.d['input_tokens']:,}／output {usage.d['output_tokens']:,} tokens；web search {usage.d['web_search_requests']} 次；估計 US${usage.usd:.2f}",
          "- 模型產生的監測清單，只作線索；進 Wattonomics 證據檔前須開原文核對。", ""]
    for cat in CATS:
        sub = [it for it in real if it.get("category") == cat]
        if not sub:
            continue
        md += [f"## {cat}（{len(sub)}）", "", "| 日期 | 玩家 | 摘要 | 園區 | MW | 金額 $B | 標記 | 來源 |", "|---|---|---|---|---|---|---|---|"]
        for it in sorted(sub, key=lambda x: str(x.get("date"))):
            md.append("| " + " | ".join(str(it.get(c) if it.get(c) is not None else "").replace("|", "／")[:60] for c in
                                         ("date", "player", "headline", "site", "mw", "amount_usd_bn", "tag")) +
                      f" | [連結]({it.get('source_url', '')}) |")
        md.append("")
    quiet = sorted({it["player"] for it in items if it.get("category") == "無"})
    if quiet:
        md.append("本週無符合消息：" + "、".join(quiet))
    fail = sorted({it["player"] for it in items if it.get("category") == "未取得"})
    if fail:
        md.append("\n**未取得結果**：" + "、".join(fail))
    open(base + ".md", "w", encoding="utf-8").write("\n".join(md) + "\n")
    return len(real)


def main():
    tpe_today = dt.datetime.now(dt.timezone(dt.timedelta(hours=8))).date()
    if tpe_today > dt.date.fromisoformat(STOP_AFTER):
        print(f"今天 {tpe_today} 已過停止日 {STOP_AFTER}（API credit 到期），不執行。")
        return
    if not C.API_KEY:
        print("未設定 ANTHROPIC_API_KEY，結束。")
        return
    reqs = sorted(glob.glob(f"{HERE}/requests/*.json"))
    jobs = [(os.path.splitext(os.path.basename(p))[0], json.load(open(p, encoding="utf-8")), p) for p in reqs]
    if not jobs:                    # 排程或手動觸發：用環境變數參數
        env = {}
        if os.environ.get("MON_PLAYERS"):
            env["players"] = [x.strip() for x in os.environ["MON_PLAYERS"].split(",") if x.strip()]
        if os.environ.get("MON_WEEK_START"):
            env["week_start"] = os.environ["MON_WEEK_START"]
        if os.environ.get("MON_SEARCHES"):
            env["searches_per_player"] = int(os.environ["MON_SEARCHES"])
        _, ws, _ = week_range(env.get("week_start"))
        jobs = [(f"{ws}_week", env, None)]
    for name, req, path in jobs:
        stamp = time.strftime("%Y%m%dT%H%M%SZ", time.gmtime())
        usage = C.Usage(TOOL, name)
        try:
            n = run(req, name, usage)
            usage.write(f"{HERE}/out/{name}_usage_{stamp}.json", {"status": "done", "items": n})
            print(name, "消息", n, f"US${usage.usd:.2f}")
            if path:
                os.makedirs(f"{HERE}/requests/done", exist_ok=True)
                shutil.move(path, f"{HERE}/requests/done/{name}.json")
        except Exception as e:
            os.makedirs(f"{HERE}/out", exist_ok=True)
            open(f"{HERE}/out/{name}.md", "w", encoding="utf-8").write(f"# 監測失敗｜{name}\n\n{type(e).__name__}: {e}\n")
            usage.write(f"{HERE}/out/{name}_usage_{stamp}.json", {"status": "failed", "error": str(e)[:300]})
            print("失敗", name, e)


if __name__ == "__main__":
    main()
