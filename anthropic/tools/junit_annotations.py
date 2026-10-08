"""CI 失敗時把 junit.xml 的失敗訊息轉成 GitHub 註記（::error::），讓 API 讀得到失敗原因（工程類，2026-10-07）。"""
import sys
import xml.etree.ElementTree as ET
from pathlib import Path

p = Path(sys.argv[1] if len(sys.argv) > 1 else "junit.xml")
if not p.exists():
    print(f"::error::找不到 {p}")
    sys.exit(0)
n = 0
for tc in ET.parse(p).getroot().iter("testcase"):
    for bad in list(tc.findall("failure")) + list(tc.findall("error")):
        msg = ((bad.get("message") or "") + " | " + (bad.text or "")[-1500:]).replace("\n", "%0A")
        print(f"::error title={tc.get('name')}::{msg[:3000]}")
        n += 1
        if n >= 8:
            sys.exit(0)
