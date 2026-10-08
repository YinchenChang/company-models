import sys, os, json, shutil, subprocess
ROOT = os.path.dirname(os.path.abspath(__file__))  # 暫存檔與重算腳本皆相對於 repo 根目錄
from openpyxl import load_workbook
src, sel, out = sys.argv[1], int(sys.argv[2]), sys.argv[3]
ak = int(sys.argv[4]) if len(sys.argv) > 4 else None  # 選填：EV/EBITDA 錨定年度
os.makedirs(os.path.join(ROOT,'out'),exist_ok=True); p=os.path.join(ROOT,'out',f'xl17_{sel}_{ak or 1}.xlsx'); shutil.copy(src,p)
wb=load_workbook(p); ws=wb['輸入與假設']
# W3（2026-10-07）：模擬運算表（H 區情境區間）改為靜態值再重算。LibreOffice 把運算表轉成 MULTIPLE.OPERATIONS，批次重算時
# 非基準情境的部分格會殘留運算表代入時的中間值（v5.26 積極情境實測：D&A 車隊 FY27–FY28 取到基準情境值，cmp31 74 項不一致）。
# 運算表輸出（保守／基準／積極三情境的加權目標價與評等）與「情境選擇」無關，直接沿用來源檔（基準情境重算後、cmp31 已核對）的快取值；
# 其餘格照常由 LibreOffice 重算。真正的 Excel 會另外計算運算表，不受影響。
from openpyxl.worksheet.formula import DataTableFormula
from openpyxl.utils import range_boundaries
_cv = load_workbook(src, data_only=True)['輸入與假設']
for _row in (list(ws.iter_rows()) if not ak else []):  # 改錨定年度時運算表輸出會變：保留運算表（verify 只在基準情境改錨定）
    for _c in _row:
        if isinstance(_c.value, DataTableFormula):
            c1, r1, c2, r2 = range_boundaries(_c.value.ref)
            for rr in range(r1, r2 + 1):
                for cc in range(c1, c2 + 1): ws.cell(row=rr, column=cc, value=_cv.cell(row=rr, column=cc).value)
for r in range(1,ws.max_row+1):
    v=ws.cell(row=r,column=1).value
    if v and str(v).startswith('情境選擇'): ws.cell(row=r,column=3,value=sel)
    if ak and v and str(v).startswith('EV/EBITDA 錨定年度'): ws.cell(row=r,column=3,value=ak)
wb.save(p)
r=subprocess.run(['python3',os.path.join(ROOT,'scripts','recalc.py'),p,'120'],capture_output=True,text=True)
if r.returncode: sys.exit('recalc 失敗：'+r.stdout+r.stderr)
wb=load_workbook(p,data_only=True); o={}
for sh in wb.sheetnames:
    ws=wb[sh]
    for r in range(1,ws.max_row+1):
        n=ws.cell(row=r,column=1).value
        if not n: continue
        if sh=='季度追蹤': o[f"{sh}|{n}"]=[ws.cell(row=r,column=3+i).value for i in range(15)]; continue  # v4.4：季度層（6 季或焦點季 15 欄；含文字）
        if sh=='輸入與假設' and str(n).startswith(('季度', '期間指引')): o[f"{sh}|{n}"]=[ws.cell(row=r,column=3+i).value for i in range(6)]; continue  # v4.4：J 區
        if sh=='評價_可比公司': v=[ws.cell(row=r,column=2+i).value for i in range(8)]
        elif sh=='檢查_連動': v=[ws.cell(row=r,column=2).value]
        else: v=[ws.cell(row=r,column=3+i).value for i in range(5)]
        if sh=='摘要' or (sh=='輸入與假設' and str(n).startswith('共識｜')): o[f"{sh}|{n}"]=v; continue  # v4.3：一頁摘要與市場共識 I 區（含文字列、擷取日期、標記）
        if any(isinstance(x,(int,float)) for x in v): o[f"{sh}|{n}"]=v
        elif sh=='評價_DCF與目標價' and (n.startswith('目標價區間｜') or n=='DCF 權重說明'): o[f"{sh}|{n}"]=[v[0] or '']  # v4.1：文字列（判斷句等）逐字比對
json.dump(o,open(out,'w'),ensure_ascii=False); print('ok',sel,len(o))
