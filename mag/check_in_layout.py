# 已移到全 repo 共用工具 tools/in_layout/check_in_layout.py（2026-10-29）；本檔保留給 MAG 的 make_v03.sh 呼叫。
import os, runpy, sys
runpy.run_path(os.path.join(os.path.dirname(__file__), "..", "tools", "in_layout", "check_in_layout.py"), run_name="__main__")
