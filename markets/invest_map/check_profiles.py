"""AI 投資地圖檢查：(1) 公司檔案網站主文不得殘留出處標記（Andy 2026-10-25 11:36 規格）；
(2) markets/invest_map 下的文字檔不得殘留 Git 合併衝突標記（2026-10-25 README 曾殘留）。
用法：python3 markets/invest_map/check_profiles.py"""
import os, re, sys
from profiles_md import load_all
HERE = os.path.dirname(os.path.abspath(__file__))
TAG = re.compile(r"\[(本專案|外部|Verified|Interested-party|Derived|判斷|Analogy|Assumed|同上)[^\]]*\]")
bad = []
for tk, d in load_all(os.path.join(HERE, "profiles")).items():
    parts = list(d["sections"].items()) + [("一句話定位", d["oneliner"])] + [("基本資料", v) for v in d["meta"].values()]
    for sec, h in parts:
        if TAG.search(re.sub(r'data-src="[^"]*"', "", h)):
            bad.append((tk, sec))
print("主文殘留出處標記：", len(bad), bad[:10])
CONFLICT = re.compile(r"^(<{7} |={7}$|>{7} )", re.M)
conf = []
for dp, _, fs in os.walk(HERE):
    for f in fs:
        if f.endswith((".md", ".py", ".sh", ".html", ".csv", ".txt")):
            fp = os.path.join(dp, f)
            if CONFLICT.search(open(fp, encoding="utf-8", errors="ignore").read()):
                conf.append(os.path.relpath(fp, HERE))
print("合併衝突標記：", len(conf), conf[:10])
sys.exit(1 if (bad or conf) else 0)
