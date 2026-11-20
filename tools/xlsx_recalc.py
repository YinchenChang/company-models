#!/usr/bin/env python3
"""把 openpyxl 寫出的 xlsx 以 LibreOffice 重算後原地存回，讓公式格帶快取值（外部稽核優化 W2-6，2026-11-20）。

openpyxl 寫出的公式格沒有計算結果；不經重算就讀值的工具（openpyxl data_only=True、預覽器、其他 AI）會讀到空白。
個股模型的成品都是重算後的檔；本模組讓直接由腳本存進 dist/ 的活頁簿（兩本地圖活頁簿）也一樣。

用法（建置腳本存檔後呼叫）：
    import xlsx_recalc; xlsx_recalc.recalc_inplace(OUT)
或：python3 tools/xlsx_recalc.py <檔案.xlsx> [...]
檢查：python3 tools/model_governance/check_cached_values.py <檔案.xlsx>
"""
import os
import shutil
import subprocess
import sys
import tempfile


def recalc_inplace(path):
    """LibreOffice headless 載入時計算全部公式，存成含快取值的 xlsx 後覆蓋原檔。失敗時丟例外、原檔不動。"""
    path = os.path.abspath(path)
    tmp = tempfile.mkdtemp()
    try:
        subprocess.run(["soffice", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", tmp, path],
                       check=True, capture_output=True)
        out = os.path.join(tmp, os.path.basename(path))
        if not os.path.exists(out):
            raise RuntimeError(f"LibreOffice 沒有產出重算檔：{path}")
        shutil.move(out, path)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    return path


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    for p in sys.argv[1:]:
        print("重算並存回：", recalc_inplace(p))
