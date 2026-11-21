# 2026-11-20 收盤換價（工作單 px1120，Andy 2026-11-21）：建置腳本在存檔前呼叫 apply(wb, ver, note)；沿革列文字讀 px1120_note_row.txt。
# 只加版本說明：README 頁最後一列追加（固定文字，開頭 v0.，屬沿革列，不進展示層關鍵數字檢查）。不插列（html.json 以固定列號取 README）。
# 數值與公式不在此改（現價與同業報價在 build 腳本輸入處改）。
from openpyxl.styles import Alignment


def apply(wb, ver, note):
    RD = wb["README"]
    r = RD.max_row + 1
    RD.cell(r, 1, note)
    RD.cell(r, 1).alignment = Alignment(wrap_text=True, vertical="top")
    return f"{ver} 版本沿革追加 README!A{r}"
