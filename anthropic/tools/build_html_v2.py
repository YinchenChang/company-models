"""相容入口：新版面產生器已移至 tools/company_html/fundmodel_v2.py（2026-11-12，泛化給 OpenAI 等收支型個股）。"""
import runpy, sys
from pathlib import Path
if "--profile" not in sys.argv:
    sys.argv += ["--profile", "anthropic"]
runpy.run_path(str(Path(__file__).resolve().parents[2] / "tools" / "company_html" / "fundmodel_v2.py"), run_name="__main__")
