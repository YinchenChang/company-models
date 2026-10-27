# 以詞法分析只改「程式碼」中的識別字，不改字串／樣板字串的文字部分（${...} 內仍視為程式碼）。
# 用法：python3 rename_ids.py  → 依 RENAMES 對 FILES 就地改名；若舊名出現在建置用的模板函式庫片段中則中止。
#       python3 rename_ids.py --fields 舊=新 …  → 改資料欄位名稱（屬性存取 .name、物件鍵 name:、解構與變數皆改；字串文字不改）。
#       新舊名稱都不得出現在模板函式庫片段；新名稱不得已存在於 FILES（5a，2026-09-26）。
import re, sys, os
H = os.path.dirname(os.path.abspath(__file__))
RENAMES = {
 'zk':'PERIODS','Lq':'PERIOD_YEARS','Bk':'UPDATE_DATE','Vk':'SOURCE_ORDER_NOTE','H1':'ACTUAL_1H',
 'Uk':'RPO_SCHEDULED_SHARE','Wk':'RPO_BUCKET_W','Gk':'RPO_Q3ADD_W','Kk':'LEASE_CASH_ON_BAL','qk':'DEBT_AMORT',
 'Jk':'LEASE_AFTER_FY30','Yk':'LATEST_Q','Xk':'CALL_FACTS','Zk':'SCENARIOS','Qk':'DEFAULTS','DBT':'DEBT_TOOLS',
 'LSE':'LEASE_FACTS','cM':'VAL_DEFAULTS','oM':'PERIOD_LABELS','sM':'HIST_PL',
 'dA':'runFunding','TM':'runValuation','rvQ':'reverseDcf','pA':'sensitivities','oA':'rpoBridge',
 'evGrid':'evAnchorGrid','gM':'dcfValue','vM':'evEbitdaLeg','xM':'blendCall','hM':'forwardPL',
 'pM':'shareCount','SM':'dcfGridWaccG','yM':'BLEND_W',
}
FILES = ['segA.js','segB.js','segC.js','segD.js','segE.js','tail.js','rm_andy.js','cmp31.js']

def split_code(s):
    """回傳 [(is_code, text)]；處理 // /* */ ' " ` 與 ${ } 巢狀。"""
    out, i, n, buf, stack = [], 0, len(s), [], []  # stack: 'tpl' 表示在樣板字串的 ${ } 內
    mode = 'code'; depth = []
    def flush(kind):
        if buf: out.append((kind, ''.join(buf))); buf.clear()
    while i < n:
        c = s[i]
        if mode == 'code':
            if c == '/' and s[i+1:i+2] == '/':
                flush(True); j = s.find('\n', i); j = n if j < 0 else j; out.append((False, s[i:j])); i = j; continue
            if c == '/' and s[i+1:i+2] == '*':
                flush(True); j = s.find('*/', i) + 2; out.append((False, s[i:j])); i = j; continue
            if c in '\'"':
                flush(True); j = i + 1
                while s[j] != c:
                    j += 2 if s[j] == '\\' else 1
                out.append((False, s[i:j+1])); i = j + 1; continue
            if c == '`':
                flush(True); mode = 'tpl'; buf.append(c); i += 1; continue
            if c == '{' and depth: depth[-1] += 1
            if c == '}' and depth:
                if depth[-1] == 0:
                    depth.pop(); flush(True); mode = 'tpl'; buf.append(c); i += 1; continue
                depth[-1] -= 1
            buf.append(c); i += 1; continue
        # 樣板字串文字
        if c == '\\':
            buf.append(s[i:i+2]); i += 2; continue
        if c == '`':
            buf.append(c); flush(False); mode = 'code'; i += 1; continue
        if c == '$' and s[i+1:i+2] == '{':
            buf.append('${'); flush(False); depth.append(0); mode = 'code'; i += 2; continue
        buf.append(c); i += 1
    flush(mode == 'code')
    return out

def rename(s, mp):
    pat = re.compile(r'(?<![\w$])(?:(?<=\.\.\.)|(?<!\.))(' +  # 允許展開運算子 ...name，排除屬性存取 .name
         '|'.join(sorted(map(re.escape, mp), key=len, reverse=True)) + r')(?![\w$])')
    res = []
    for is_code, t in split_code(s):
        if is_code:
            # 物件屬性鍵（name:）不改，避免改到資料欄位
            t = pat.sub(lambda m: m.group(1) if re.match(r'\s*:(?!:)', t[m.end():m.end()+3]) and not t[:m.start()].rstrip().endswith('?') else mp[m.group(1)], t)
        res.append(t)
    return ''.join(res)

def rename_fields(s, mp):
    """資料欄位改名：程式碼中的完整識別字一律替換（含 .name 與 name:），字串與樣板字串的文字部分不動。"""
    pat = re.compile(r'(?<![\w$])(' + '|'.join(sorted(map(re.escape, mp), key=len, reverse=True)) + r')(?![\w$])')
    return ''.join(pat.sub(lambda m: mp[m.group(1)], t) if k else t for k, t in split_code(s))


def lib_code():
    L = open(os.path.join(H, 'app_pretty.js'), encoding='utf-8').read().split('\n')
    return ''.join(t for k, t in split_code('\n'.join(L[685:2518] + L[2940:2976] + L[3649:3751])) if k)


if __name__ == '__main__' and sys.argv[1:2] == ['--fields']:
    mp = dict(a.split('=') for a in sys.argv[2:])
    code = lib_code()
    bad = [n for pair in mp.items() for n in pair if re.search(r'(?<![\w$])' + re.escape(n) + r'(?![\w$])', code)]
    if bad: sys.exit(f'模板函式庫片段使用了這些名稱，中止：{bad}')
    clash = [nw for nw in mp.values() if any(re.search(r'(?<![\w$])' + re.escape(nw) + r'(?![\w$])', open(os.path.join(H, f), encoding='utf-8').read()) for f in FILES)]
    if clash: sys.exit(f'新名稱已存在，中止：{clash}')
    for f in FILES:
        p = os.path.join(H, f); s = open(p, encoding='utf-8').read()
        assert ''.join(t for _, t in split_code(s)) == s, f'lexer 重組失敗 {f}'
        open(p, 'w', encoding='utf-8').write(rename_fields(s, mp))
    print('renamed fields', mp); sys.exit(0)

if __name__ == '__main__':
    L = open(os.path.join(H, 'app_pretty.js'), encoding='utf-8').read().split('\n')
    chunks = '\n'.join(L[685:2518] + L[2940:2976] + L[3649:3751])
    code_chunks = ''.join(t for k, t in split_code(chunks) if k)
    bad = [o for o in RENAMES if re.search(r'(?<![\w$])(?:(?<=\.\.\.)|(?<!\.))' + re.escape(o) + r'(?![\w$])', code_chunks)]
    clash = [nw for nw in RENAMES.values() if any(re.search(r'(?<![\w$])' + nw + r'(?![\w$])', open(os.path.join(H, f), encoding='utf-8').read()) for f in FILES if f != 'segA.js') and nw not in ('PERIODS','PERIOD_LABELS','PERIOD_YEARS','UPDATE_DATE','SOURCE_ORDER_NOTE','ACTUAL_1H','RPO_SCHEDULED_SHARE','RPO_BUCKET_W','RPO_Q3ADD_W','LEASE_CASH_ON_BAL','DEBT_AMORT','LEASE_AFTER_FY30','LATEST_Q','CALL_FACTS','SCENARIOS','DEBT_TOOLS','LEASE_FACTS','DEFAULTS')]
    if bad: print('模板片段使用了這些舊名，不改：', bad); [RENAMES.pop(b) for b in bad]
    if clash: print('新名稱已存在，中止：', clash); sys.exit(1)
    for f in FILES:
        p = os.path.join(H, f); s = open(p, encoding='utf-8').read()
        # 自檢：切分後重組必須與原文相同
        assert ''.join(t for _, t in split_code(s)) == s, f'lexer 重組失敗 {f}'
        open(p, 'w', encoding='utf-8').write(rename(s, RENAMES))
    print('renamed', len(RENAMES), 'names in', len(FILES), 'files')
