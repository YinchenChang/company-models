"""Siliconomics 展示層：把重算後的 xlsx（siliconomics/model/CURRENT）轉成單頁自含 HTML。

用法：python3 siliconomics/tools/build_html.py [xlsx] <輸出.html>
  xlsx 省略時讀 siliconomics/model/CURRENT 指向的檔案。
規則：
  - 數字一律取 xlsx 的重算值（data_only），格式依儲存格 number_format；本程式不做任何計算。
  - 每個工作表一個分頁，保留列順序；藍字（輸入）、淡黃底（關鍵輸出）、亮黃底（待決定）沿用 xlsx 顏色語意。
  - 每格以 title 標註來源儲存格（例如 Rack!C6），方便回 Excel 核對。
  - 發布：markets/invest_map/publish_pages.sh 在發布時呼叫本程式，輸出到 GitHub Pages 的 siliconomics/index.html。
"""
import html, os, re, sys
import openpyxl
from openpyxl.utils import get_column_letter

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
MODEL = os.path.join(ROOT, "siliconomics", "model")

# 分頁順序與中文說明（未列出的工作表接在最後）
TABS = [
    ("README", "導讀", "版本沿革、讀法、顏色與標記規則"),
    ("Scen", "情境", "A 順利／B 平平／C 不如預期的定義"),
    ("Deploy", "部署 GW", "AI IT 資本支出與部署 GW（三情境）"),
    ("ShareTbl", "份額表", "各子產業廠商份額與加總檢查"),
    ("ABF", "ABF 供需", "ABF 載板需求、供給與缺口"),
    ("PGW", "每 GW PCB／CCL", "每 GW 高速 PCB 與 CCL 金額"),
    ("Rack", "機架層市場", "伺服器組裝、商用 GPU 系統、AI 電源、液冷市場"),
    ("Equip", "設備", "SEMI WFE、SoC 測試機市場、混合鍵合"),
    ("Path", "記憶體價格路徑", "DRAM／NAND 季價路徑"),
    ("Mkt", "加速器與 HBM 市場", "加速器顆數層與 HBM 市場"),
    ("Units", "平台顆數", "各平台加速器顆數"),
    ("TK_In", "Tokenomics 鏡像", "取自算力收支 Tokenomics 的名稱與值"),
    ("In", "輸入", "全部行業輸入：基準、區間、標記、來源"),
    ("Interface", "下游取數", "SIF_ 具名範圍（各公司模型讀這裡）"),
    ("Conflicts", "衝突登錄", "跨模型口徑衝突與處理狀態"),
    ("Decisions", "決策", "已定案的方法決定"),
    ("Checks", "檢查", "交叉檢查（ERROR 須為 0）"),
    ("DB_Evidence", "證據登錄", "外部證據送件"),
]
BIG = {"In", "Interface"}   # 列數多的頁：預設顯示搜尋框提示


def fmt(v, nf):
    """依 Excel number_format 轉成顯示字串（只處理本活頁簿用到的格式）。"""
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, (int, float)):
        nf = (nf or "General").split(";")[0]
        if "%" in nf:
            m = re.search(r"0\.(0+)%", nf)
            d = len(m.group(1)) if m else 0
            return f"{v * 100:,.{d}f}%"
        if nf == "@":
            return str(v)
        m = re.search(r"0\.(0+)", nf)
        if m:
            return f"{v:,.{len(m.group(1))}f}"
        if re.search(r"[#0]", nf):
            return f"{v:,.0f}"
        # General：整數照印，小數最多 4 位有效位數
        if float(v).is_integer():
            return f"{int(v):,}" if abs(v) >= 10000 else str(int(v))
        a = abs(v)
        d = 2 if a >= 100 else 3 if a >= 1 else 4
        return f"{v:,.{d}f}".rstrip("0").rstrip(".")
    if hasattr(v, "strftime"):
        return v.strftime("%Y-%m-%d")
    return str(v)


def rgb(c):
    try:
        x = c.rgb
        return x[-6:].upper() if isinstance(x, str) and len(x) >= 6 else None
    except Exception:
        return None


def cls_of(cell):
    out = []
    f = cell.font
    fc = rgb(f.color) if f and f.color is not None else None
    if fc in ("0000FF", "0000CC", "1F4E79", "0070C0", "0432FF"):
        out.append("in")
    elif fc in ("008000", "00B050", "007A33"):
        out.append("ln")
    if f and f.bold:
        out.append("b")
    fill = cell.fill
    if fill is not None and fill.fill_type == "solid":
        bg = rgb(fill.fgColor)
        if bg in ("FFFF00",):
            out.append("hot")
        elif bg in ("DDEBF7",):
            out.append("th")
        elif bg and bg.startswith("FF") and bg not in ("FFFFFF",):
            out.append("key")
        elif bg and bg not in ("FFFFFF", "000000"):
            out.append("key")
    if isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool):
        out.append("n")
    v = cell.value
    if isinstance(v, str):
        if v in ("ERROR",):
            out.append("err")
        elif v in ("WARN",):
            out.append("warn")
        elif v == "OK":
            out.append("ok")
    return " ".join(out)


def sheet_html(ws):
    rows = list(ws.iter_rows())
    # 裁掉尾端空列與右側空欄
    maxc = 0
    keep = []
    for row in rows:
        last = 0
        for c in row:
            if c.value not in (None, ""):
                last = c.column
        keep.append(last)
        maxc = max(maxc, last)
    while keep and keep[-1] == 0:
        keep.pop(); rows.pop()
    if not rows:
        return "<p class='empty'>（空白頁）</p>"
    out = ["<div class='tw'><table>"]
    blank_run = 0
    for row, last in zip(rows, keep):
        if last == 0:
            blank_run += 1
            if blank_run == 1:
                out.append("<tr class='sp'><td colspan='%d'></td></tr>" % maxc)
            continue
        blank_run = 0
        r = row[0].row
        # 只有 A 欄有字的列＝標題或說明列：跨欄顯示
        if last == 1 and isinstance(row[0].value, str):
            c = row[0]
            k = "hd" if (c.font and c.font.bold) or r == 1 else "note"
            out.append(f"<tr class='{k}'><td colspan='{maxc}' title='{ws.title}!A{r}'>{html.escape(fmt(c.value, c.number_format))}</td></tr>")
            continue
        tds = []
        for c in row[:maxc]:
            ref = f"{ws.title}!{get_column_letter(c.column)}{r}"
            txt = fmt(c.value, c.number_format)
            k = cls_of(c)
            long = isinstance(c.value, str) and len(c.value) > 60
            if long:
                k = (k + (" wr" if c.column == 1 else " lg")).strip()
            tds.append(f"<td class='{k}' title='{ref}'>{html.escape(txt)}</td>" if txt else "<td></td>")
        out.append("<tr>" + "".join(tds) + "</tr>")
    out.append("</table></div>")
    return "\n".join(out)


def main():
    args = sys.argv[1:]
    if len(args) == 1:
        cur = open(os.path.join(MODEL, "CURRENT"), encoding="utf-8").read().strip()
        src, dst = os.path.join(MODEL, cur), args[0]
    elif len(args) == 2:
        src, dst = args
    else:
        sys.exit(__doc__)
    wb = openpyxl.load_workbook(src, data_only=True)
    title_line = str(wb["README"]["A1"].value or os.path.basename(src))
    m = re.search(r"v\d+(\.\d+)*", os.path.basename(src))
    ver = m.group(0) if m else ""
    names = [t[0] for t in TABS if t[0] in wb.sheetnames] + [n for n in wb.sheetnames if n not in {t[0] for t in TABS}]
    meta = {t[0]: t for t in TABS}
    nav, panes = [], []
    for i, n in enumerate(names):
        _, zh, desc = meta.get(n, (n, n, ""))
        nav.append(f"<button class='tab{' on' if i == 0 else ''}' data-t='p{i}' title='{html.escape(desc)}'>{html.escape(zh)}<small>{html.escape(n)}</small></button>")
        hint = "（本頁列數多，可用上方搜尋）" if n in BIG else ""
        panes.append(f"<section id='p{i}' class='pane{' on' if i == 0 else ''}'><h2>{html.escape(zh)} <span>{html.escape(n)}</span></h2>"
                     f"<p class='desc'>{html.escape(desc)}{hint}</p>{sheet_html(wb[n])}</section>")
    page = TEMPLATE.replace("{{TITLE}}", html.escape(f"Siliconomics {ver}")) \
        .replace("{{HEAD}}", html.escape(title_line)) \
        .replace("{{SRC}}", html.escape("siliconomics/model/" + os.path.basename(src))) \
        .replace("{{NAV}}", "\n".join(nav)).replace("{{PANES}}", "\n".join(panes))
    os.makedirs(os.path.dirname(os.path.abspath(dst)), exist_ok=True)
    open(dst, "w", encoding="utf-8").write(page)
    print(f"{dst}（{len(page):,} bytes，{len(names)} 頁）")


TEMPLATE = """<!doctype html>
<html lang="zh-Hant"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>{{TITLE}}</title>
<style>
:root{--bg:#fbfaf7;--fg:#1d1f24;--mut:#6b6f78;--line:#e3e0d8;--card:#fff;--in:#1f4fbf;--ln:#1c7a3e;--key:#fff6cf;--hot:#ffe14d;--acc:#8a5a00;--err:#b3261e;--ok:#1c7a3e}
@media (prefers-color-scheme:dark){:root{--bg:#16171a;--fg:#e8e6e1;--mut:#9a9ca3;--line:#2c2e33;--card:#1d1f23;--in:#8fb2ff;--ln:#7fd39b;--key:#3a3420;--hot:#6b5a10;--acc:#e0b860;--err:#ff8a80;--ok:#7fd39b}}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--fg);font:14px/1.5 -apple-system,"Segoe UI","Noto Sans TC","Microsoft JhengHei",sans-serif}
header{padding:20px 16px 8px;max-width:1400px;margin:0 auto}
h1{font-size:22px;margin:0 0 4px}header p{margin:2px 0;color:var(--mut);font-size:13px}
.bar{position:sticky;top:0;z-index:5;background:var(--bg);border-bottom:1px solid var(--line);padding:8px 16px}
.bar .in{max-width:1400px;margin:0 auto}
nav{display:flex;gap:6px;overflow-x:auto;padding-bottom:6px;scrollbar-width:thin}
.tab{flex:0 0 auto;border:1px solid var(--line);background:var(--card);color:var(--fg);border-radius:8px;padding:5px 10px;cursor:pointer;font:inherit;text-align:left}
.tab small{display:block;color:var(--mut);font-size:11px}.tab.on{border-color:var(--acc);box-shadow:inset 0 -2px 0 var(--acc)}
#q{width:100%;max-width:420px;padding:6px 10px;border:1px solid var(--line);border-radius:8px;background:var(--card);color:var(--fg);font:inherit}
main{max-width:1400px;margin:0 auto;padding:12px 16px 40px}
.pane{display:none}.pane.on{display:block}
h2{font-size:18px;margin:6px 0 2px}h2 span{color:var(--mut);font-size:13px;font-weight:400}
.desc{color:var(--mut);margin:0 0 10px;font-size:13px}
.tw{overflow-x:auto;border:1px solid var(--line);border-radius:8px;background:var(--card)}
table{border-collapse:collapse;width:100%;font-size:13px}
td{border-bottom:1px solid var(--line);padding:4px 8px;vertical-align:top;white-space:nowrap}
td.wr{white-space:normal;min-width:240px}
tr:not(.hd):not(.note) td:first-child{position:sticky;left:0;background:var(--card);z-index:1}
@media (max-width:640px){td.wr,tr:not(.hd):not(.note) td:first-child{min-width:150px;max-width:170px;white-space:normal}}
td.lg{white-space:normal;min-width:320px;max-width:560px;color:var(--mut)}
tr.hd td{font-weight:600;background:color-mix(in srgb,var(--acc) 8%,var(--card));white-space:normal}
tr.note td{color:var(--mut);white-space:normal}
tr.sp td{border:0;padding:3px}
td.n{text-align:right;font-variant-numeric:tabular-nums}
td.in{color:var(--in)}td.ln{color:var(--ln)}td.b{font-weight:600}
td.key{background:var(--key)}td.th{background:color-mix(in srgb,var(--in) 10%,var(--card));font-weight:600}td.hot{background:var(--hot)}
td.err{color:var(--err);font-weight:700}td.warn{color:var(--acc);font-weight:600}td.ok{color:var(--ok)}
tr.hide{display:none}
.legend{display:flex;flex-wrap:wrap;gap:12px;font-size:12px;color:var(--mut);margin-top:6px}
.legend i{display:inline-block;width:12px;height:12px;border-radius:2px;vertical-align:-2px;margin-right:4px;border:1px solid var(--line)}
footer{max-width:1400px;margin:0 auto;padding:0 16px 30px;color:var(--mut);font-size:12px}
</style></head><body>
<header><h1>{{HEAD}}</h1>
<p>AI 半導體各子產業的行業變數基準庫（唯一事實來源）。本頁為 Excel 的展示層，所有數字取自重算後的活頁簿，滑鼠停在數字上可看到對應儲存格。</p>
<p>來源檔：{{SRC}}（repo YinchenChang/company-models）。研究整理，非投資建議。</p>
<div class="legend"><span><i style="background:transparent;border-color:var(--in)"></i><b style="color:var(--in)">藍字</b>＝輸入</span><span><b style="color:var(--ln)">綠字</b>＝跨頁或外部連結</span><span><i style="background:var(--key)"></i>關鍵輸出</span><span><i style="background:var(--hot)"></i>待決定</span></div>
</header>
<div class="bar"><div class="in"><nav>{{NAV}}</nav><input id="q" type="search" placeholder="搜尋本頁（例如：HBM、ABF、SIF_、Ibiden）"></div></div>
<main>{{PANES}}</main>
<footer>標記：Verified 已驗證／Interested-party 利害關係方／Analogy 類比／Assumed 假設／Derived 推得。Analogy 與 Assumed 一律給區間。</footer>
<script>
const tabs=[...document.querySelectorAll('.tab')],q=document.getElementById('q');
function show(id){tabs.forEach(t=>t.classList.toggle('on',t.dataset.t===id));document.querySelectorAll('.pane').forEach(p=>p.classList.toggle('on',p.id===id));try{history.replaceState(null,'','#'+id)}catch(e){}filt()}
tabs.forEach(t=>t.addEventListener('click',()=>show(t.dataset.t)));
function filt(){const s=q.value.trim().toLowerCase(),p=document.querySelector('.pane.on');if(!p)return;p.querySelectorAll('tr').forEach(tr=>{tr.classList.toggle('hide',!!s&&!tr.textContent.toLowerCase().includes(s))})}
q.addEventListener('input',filt);
if(location.hash&&document.getElementById(location.hash.slice(1)))show(location.hash.slice(1));
</script></body></html>
"""

if __name__ == "__main__":
    main()
