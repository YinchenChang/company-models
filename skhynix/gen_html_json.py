# 由 html_src.json（原 html.json，內容不動）產生 skhynix/html.json：套用個股 HTML 版面規格 v1.0（docs/plan/個股HTML版面規格.md）。
# 用法：python3 skhynix/gen_html_json.py <成品 xlsx>（勿手改 html.json）
# 記憶體三家共用的調整在 tools/memory_model/html_layout.py（轉換器 tools/company_html/relayout.py）。
import os, sys
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "tools", "memory_model"))
import html_layout  # noqa: E402
html_layout.build(HERE, sys.argv[1], nm_block=False, pfmt="n0", v13=True)   # v0.7.6 起：版面 v1.3
