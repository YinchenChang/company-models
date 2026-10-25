"""公司檔案檢查：網站主文不得殘留出處標記（Andy 2026-10-25 11:36 規格）。用法：python3 markets/invest_map/check_profiles.py"""
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
sys.exit(1 if bad else 0)
