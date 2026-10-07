# 用法：python3 build_html_portable.py <版本> <輸出路徑> <日期> [模板 HTML]（模板省略時用 repo 內的 docs/template_v3_3.html；相對路徑以目前目錄為準）
# 模板可用任一版已發布的 HTML（例如 20260923_CoreWeave收支模型_v3_3.html），只取其中的 React／Recharts 函式庫部分。
import sys, os
H = os.path.dirname(os.path.abspath(__file__))
ver, out, date = sys.argv[1], sys.argv[2], sys.argv[3]
tpl = sys.argv[4] if len(sys.argv) > 4 else os.path.join(H, 'docs', 'template_v3_3.html')
L = open(os.path.join(H, 'app_pretty.js'), encoding='utf-8').read().split('\n')
rng = lambda a, b: '\n'.join(L[a-1:b])
mid1 = rng(686, 2518); mid2 = rng(2941, 2976); mid3 = rng(3650, 3751)
mid3 = mid3.replace("var NM = [`年度假設`, `年度結果`, `收入／容量`, `成本`, `信用`, `終值`, `園區`, `敏感性`, `連動檢查`, `來源`];",
                    "var NM = [`支出假設`, `各期結果`, `收入／產能`, `電力成本`, `信用／利率`, `終值`, `站點`, `敏感性`, `連動檢查`, `來源與承諾`, `債務明細`];")
k = mid3.index('function RM('); ke = mid3.find('\nfunction ', k + 10); ke = len(mid3) if ke < 0 else ke
mid3 = mid3[:k] + open(os.path.join(H, 'rm_andy.js'), encoding='utf-8').read().rstrip('\n') + mid3[ke:]
seg = lambda n: open(os.path.join(H, n), encoding='utf-8').read()
import json as _json
sys.path.insert(0, H); import calendar_q  # v4.5：期間與日期由 company.json → calendar 推算（與 Excel 共用）
_co = calendar_q.load(H)  # 公司資料單一來源
_co.pop('asOf', None)  # 滾動檢查的季度標記只在建置時檢查（calendar_q），不注入 HTML
_co['consensus'] = _json.load(open(os.path.join(H, _co['meta']['consensusFile']), encoding='utf-8'))  # v4.3：市場共識資料檔（只讀）併入注入資料，不另設全域變數
_g = _co['consensus']['companyGuidance'].get('2026Q3') or {}  # Q3 營收指引以 company.json 為準；與共識檔不一致即停止建置（Excel 建置同一檢查；v0.1b：公司未給季度指引時兩邊皆為空）
assert (_g.get('revenueLow'), _g.get('revenueHigh')) == (_co['callFacts']['nextQRevLo'], _co['callFacts']['nextQRevHi']), 'Q3 營收指引：company.json 與共識檔不一致'
assert abs(_co['ytdActual']['adjEbitda'] - _co['ytdActual']['adjEbitdaMeta']['q1'] - _co['ytdActual']['adjEbitdaMeta']['q2']) < 1e-9, '1H 調整後 EBITDA ≠ Q1＋Q2'
import subprocess as _sp  # v4.4：季度加總＝年度、指引一致性、超過門檻的差距都有原因；不符即停止建置
_ck = _sp.run(['node', os.path.join(H, 'scripts', 'check_quarterly.js'), H], capture_output=True, text=True)
# CRWV_SKIP_QCHECK：只供 scripts/test_rolling.py 的暫存副本使用（只改日曆、不改數字，季度一致性必然不符）
if _ck.returncode and not os.environ.get('CRWV_SKIP_QCHECK'): sys.exit('check_quarterly 失敗：\n' + _ck.stdout + _ck.stderr)
# 注意：模板函式庫在插入點前是一個未結束的 var 宣告鏈（…,Rk=…,），所以第一個敘述必須是「名稱 = 值」接續該鏈
app = '\n'.join(['COMPANY_DATA = ' + _json.dumps(_co, ensure_ascii=False) + ';', seg('segA.js'), mid1, seg('segB.js'), mid2, seg('segC.js'), mid3, seg('segD.js'), seg('tail.js'), seg('segE.js')])
s = open(tpl, encoding='utf-8').read()
i = -1
for _mk in ('COMPANY_DATA = {"meta"', 'zk=[`FY27`', 'zk = [`FY26`'):  # v4.0 起以 var CO 為起點；相容舊模板
    i = s.find(_mk)
    if i >= 0: break
j = s.find('</script>', i)
html = s[:i] + app + s[j:]
import re
_nm, _th = _co['meta']['company'], _co['texts']['thesis']  # v0.1b：公司名稱與命題讀 company.json
html = re.sub(r'<!-- .*?-->', f'<!-- {_th} · {_nm} 收支模型 v{ver}（更新 {date}）· 雙擊以 Chrome / Edge 開啟，不需安裝或連網 -->', html, count=1)
html = re.sub(r'<title>.*?</title>', f'<title>{_th} · {_nm} 收支模型 v{ver} · {date} · 離線版</title>', html, count=1)
tag = ver.replace('.', '')
html = html.replace('`v`, `1.0`', f'`v`, `{ver}`').replace('crwv_annual_v10', f'crwv_annual_v{tag}').replace('crwv_sites_v10', f'crwv_sites_v{tag}').replace('crwv_model_v10', f'crwv_model_v{tag}')
open(out, 'w', encoding='utf-8').write(html)
print('built', len(html))
