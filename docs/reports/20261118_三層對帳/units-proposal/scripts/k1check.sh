#!/bin/bash
# cascade 後 K1 自舉：以情境複本現行活頁簿為底稿重建 Siliconomics，LibreOffice 重算，回報 SGOV_Errors 與錯誤格
N=$1; D=/tmp/pp_$N; cd $D
F=$(cat siliconomics/model/CURRENT); V=${F#20261118_Siliconomics_}; V=${V%.xlsx}
mkdir -p build/k1/raw build/k1/out; cp siliconomics/model/$F build/k1/base_$F
python3 siliconomics/builder/build.py --base build/k1/base_$F --out build/k1/raw/$F --version $V --date 20261118 > build/k1/build.log 2>&1
soffice -env:UserInstallation=file:///tmp/lo_k1_$N --headless --calc --convert-to xlsx:"Calc MS Excel 2007 XML" --outdir build/k1/out build/k1/raw/$F > /dev/null 2>&1
python3 - "$D/build/k1/out/$F" "$D/siliconomics/model/$F" <<'P'
import sys, openpyxl
a = openpyxl.load_workbook(sys.argv[1], data_only=True)
sh, ref = list(a.defined_names["SGOV_Errors"].destinations)[0]
print("SGOV_Errors after K1 rebuild:", a[sh][ref.replace("$", "")].value)
P
python3 /home/claude/wt/proposal/tools/xlsx_diff.py build/k1/out/$F siliconomics/model/$F --tol 1e-9 --limit 30 2>&1 | grep -E "錯誤格|差異" | head -n 4
