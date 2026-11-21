"""抓 Yahoo Finance 日線收盤（未還原股利；交易所當地日期），供 px 工作單使用。
用法：python3 tools/px_closes/fetch_closes.py <起日> <迄日> <輸出 csv> [額外代號…]
代號清單＝markets/invest_map/data/closes_*.csv 既有代號 ∪ 額外代號；幣別沿用既有檔。
在 GitHub Actions 執行（雲端工作環境連不到 Yahoo）；用 yfinance（直接打 chart API 在 runner 上會被限流）。
抓不到的代號列在 <輸出>.missing.txt，不臆造；期間內的分割列在 <輸出>.splits.txt。"""
import csv, glob, sys, time, datetime as dt
import yfinance as yf

start, end, out = sys.argv[1], sys.argv[2], sys.argv[3]
extra = sys.argv[4:]
cur = {}
for f in sorted(glob.glob("markets/invest_map/data/closes_*.csv")):
    if f == out: continue
    for r in csv.DictReader(open(f, encoding="utf-8")):
        cur.setdefault(r["symbol"], r["currency"])
syms = sorted(set(cur) | set(extra))
end1 = (dt.datetime.strptime(end, "%Y-%m-%d") + dt.timedelta(days=1)).strftime("%Y-%m-%d")
rows, got, splits, meta = [], set(), [], []
def grab(batch):
    df = yf.download(batch, start=start, end=end1, interval="1d", auto_adjust=False, actions=True,
                     group_by="ticker", threads=False, progress=False)
    for s in batch:
        try:
            sub = df[s] if len(batch) > 1 else df
            if hasattr(sub.columns, "nlevels") and sub.columns.nlevels > 1: sub = sub[s]
            for d, c in sub["Close"].dropna().items():
                rows.append((s, cur.get(s, ""), d.strftime("%Y-%m-%d"), repr(round(float(c), 6)))); got.add(s)
            if "Stock Splits" in sub:
                for d, v in sub["Stock Splits"].dropna().items():
                    if float(v) != 0: splits.append("%s,%s,%s" % (s, d.strftime("%Y-%m-%d"), v))
        except Exception as e:
            print("ERR", s, repr(e)[:120])
for i in range(0, len(syms), 25):
    grab(syms[i:i + 25]); time.sleep(2)
for rnd in range(2):
    miss = [s for s in syms if s not in got]
    if not miss: break
    time.sleep(20)
    for s in miss: grab([s]); time.sleep(1)
# 個別代號最後一日落後於全體最後一日時（批次下載偶爾漏最後一根；分割日亦然），改用 Ticker.history 補
last = {}
for r in rows: last[r[0]] = max(last.get(r[0], ""), r[2])
top = max(last.values())
for s in [s for s in syms if last.get(s, "") < top]:
    try:
        h = yf.Ticker(s).history(start=start, end=end1, interval="1d", auto_adjust=False, actions=True)
        for d, c in h["Close"].dropna().items():
            rows.append((s, cur.get(s, ""), d.strftime("%Y-%m-%d"), repr(round(float(c), 6)))); got.add(s)
        print("refetch", s, h["Close"].dropna().index[-1])
        # 日線仍缺最後一日時，取報價中繼資料的最近成交價（regularMarketPrice）與其當地日期；另記在 .meta.txt
        t = yf.Ticker(s); t.history(period="5d"); m = t.history_metadata
        import zoneinfo
        d = dt.datetime.fromtimestamp(m["regularMarketTime"], zoneinfo.ZoneInfo(m["exchangeTimezoneName"])).strftime("%Y-%m-%d")
        print("meta", s, d, m.get("regularMarketPrice"), m.get("chartPreviousClose"))
        if h["Close"].dropna().index[-1].strftime("%Y-%m-%d") < d <= end and m.get("regularMarketPrice"):
            rows.append((s, cur.get(s, ""), d, repr(round(float(m["regularMarketPrice"]), 6)))); meta.append("%s,%s,%s" % (s, d, m["regularMarketPrice"]))
    except Exception as e:
        print("ERR refetch", s, repr(e)[:120])
    time.sleep(1)
missing = [s for s in syms if s not in got]
rows = sorted(set(rows))
with open(out, "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f); w.writerow(["symbol", "currency", "date", "close"]); w.writerows(rows)
open(out + ".missing.txt", "w").write("\n".join(missing) + "\n")
open(out + ".meta.txt", "w").write("\n".join(meta) + "\n")
open(out + ".splits.txt", "w").write("\n".join(splits) + "\n")
print("symbols", len(syms), "rows", len(rows), "missing", missing, "splits", splits)
