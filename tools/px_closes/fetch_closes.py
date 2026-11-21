"""抓 Yahoo Finance 日線收盤（未還原股利；交易所當地日期），供 px 工作單使用。
用法：python3 tools/px_closes/fetch_closes.py <起日> <迄日> <輸出 csv> [額外代號…]
代號清單＝markets/invest_map/data/closes_*.csv 既有代號 ∪ 額外代號；幣別沿用既有檔。
在 GitHub Actions 執行（雲端工作環境連不到 Yahoo）。抓不到的代號列在 <輸出>.missing.txt，不臆造。"""
import csv, glob, json, sys, time, datetime as dt, urllib.request, urllib.parse

start, end, out = sys.argv[1], sys.argv[2], sys.argv[3]
extra = sys.argv[4:]
cur = {}
for f in sorted(glob.glob("markets/invest_map/data/closes_*.csv")):
    for r in csv.DictReader(open(f, encoding="utf-8")):
        cur.setdefault(r["symbol"], r["currency"])
syms = sorted(set(cur) | set(extra))
p1 = int(dt.datetime.strptime(start, "%Y-%m-%d").replace(tzinfo=dt.timezone.utc).timestamp())
p2 = int((dt.datetime.strptime(end, "%Y-%m-%d") + dt.timedelta(days=1)).replace(tzinfo=dt.timezone.utc).timestamp())
rows, missing, splits = [], [], []
for s in syms:
    url = ("https://query1.finance.yahoo.com/v8/finance/chart/%s?period1=%d&period2=%d&interval=1d&events=div%%2Csplits"
           % (urllib.parse.quote(s), p1, p2))
    j = None
    for k in range(4):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/126 Safari/537.36"})
            j = json.load(urllib.request.urlopen(req, timeout=30)); break
        except Exception as e:
            err = repr(e); time.sleep(3 * (k + 1))
    try:
        res = j["chart"]["result"][0]
        off = res["meta"].get("gmtoffset", 0)
        ccy = res["meta"].get("currency") or cur.get(s, "")
        ts, cl = res["timestamp"], res["indicators"]["quote"][0]["close"]
        n = 0
        for t, c in zip(ts, cl):
            if c is None: continue
            d = dt.datetime.fromtimestamp(t + off, dt.timezone.utc).strftime("%Y-%m-%d")
            if start <= d <= end:
                rows.append((s, cur.get(s) or ccy, d, repr(round(c, 6)))); n += 1
        for ev in (res.get("events", {}).get("splits", {}) or {}).values():
            splits.append("%s,%s,%s" % (s, dt.datetime.fromtimestamp(ev["date"] + off, dt.timezone.utc).strftime("%Y-%m-%d"), ev.get("splitRatio")))
        if n == 0: missing.append(s + ",no rows")
    except Exception as e:
        missing.append("%s,%s" % (s, (j or {}).get("chart", {}).get("error") if j else err))
    time.sleep(0.4)
with open(out, "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f); w.writerow(["symbol", "currency", "date", "close"]); w.writerows(rows)
open(out + ".missing.txt", "w").write("\n".join(missing) + "\n")
open(out + ".splits.txt", "w").write("\n".join(splits) + "\n")
print("symbols", len(syms), "rows", len(rows), "missing", len(missing)); print("\n".join(missing)); print("splits:", splits)
