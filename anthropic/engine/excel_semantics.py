"""pycel 外掛：補齊與 Excel 語意不一致的函數（pycel 以 plugins 優先於內建函數載入）。

只補「語意」，不新增任何計算機制。目前一項：
- TEXT(數值, "0" / "0.0" / "0.00" …)：Excel 以 15 位有效數字的十進位表示做「四捨五入（遠離零）」，
  pycel 內建以二進位浮點與銀行家進位，於 0.25→"0.3"、2.675→"2.68" 等平手值不同。
"""
import re
from decimal import ROUND_HALF_UP, Decimal

from pycel.lib.function_helpers import excel_helper
from pycel.lib.text import text as _pycel_text

_PLAIN_FMT = re.compile(r"^0(?:\.(0+))?$")


@excel_helper(cse_params=0, str_params=1)
def text(text_value, value_format):
    m = _PLAIN_FMT.match(value_format) if isinstance(value_format, str) else None
    if m and isinstance(text_value, (int, float)) and not isinstance(text_value, bool):
        places = len(m.group(1) or "")
        d = Decimal(format(text_value, ".15g"))                 # Excel 的 15 位有效數字
        q = d.quantize(Decimal(1).scaleb(-places), rounding=ROUND_HALF_UP)
        s = format(q, "f")
        return "0" + s[1:] if s.startswith("-0") and q == 0 else s  # Excel 不輸出 "-0"
    return _pycel_text(text_value, value_format)   # 其他格式沿用 pycel 內建
