#!/bin/bash
# 情境建置：在 /tmp 複本重建 Siliconomics（兩遍：第二遍讀新的 Recon 快照），可選跑協調橋接與證據比例診斷。
# 用法：run_sil.sh <情境名> <來源樹> <edits.json 或 none> <full|light> [pc28=0|1]
set -eu
N=$1; SRC=$2; ED=$3; MODE=$4; PC=${5:-0}
D=/tmp/pp_$N; LOG=/home/claude/out/proposal/runs/$N.log
rm -rf $D; mkdir -p $D
if [ $MODE = full ]; then cp -r $SRC/. $D/; else for x in siliconomics markets wattonomics tools tokenomics_snap docs; do [ -e $SRC/$x ] && cp -r $SRC/$x $D/; done; fi
cd $D; mkdir -p build/raw
export PYTHONDONTWRITEBYTECODE=1
LO="-env:UserInstallation=file:///tmp/lo_$N"
CUR=siliconomics/model/$(cat siliconomics/model/CURRENT)
V=v0.21-trial-$N; F=20261118_Siliconomics_$V.xlsx
if [ "$ED" != none ]; then python3 siliconomics/tools/apply_inputs.py $CUR $ED build/in.xlsx > build/apply.log; else cp $CUR build/in.xlsx; fi
if [ "$PC" = 0 ]; then rm -f siliconomics/builder/seed/rows_added_v0.21.json siliconomics/builder/seed/decisions_v0.21.json; fi
bld() { python3 siliconomics/builder/build.py --base $1 --out build/raw/$F --version $V --date 20261118 > build/build_$2.log 2>&1
        soffice $LO --headless --calc --convert-to xlsx:"Calc MS Excel 2007 XML" --outdir siliconomics/model build/raw/$F > /dev/null 2>&1; }
bld build/in.xlsx 1
git -C $D mv -f $CUR siliconomics/model/archive/ 2>/dev/null || mv $CUR siliconomics/model/archive/
echo $F > siliconomics/model/CURRENT
python3 siliconomics/tools/r4_bridge.py build/r4_raw.xlsx > build/r4.log 2>&1
soffice $LO --headless --calc --convert-to xlsx:"Calc MS Excel 2007 XML" --outdir build/r4 build/r4_raw.xlsx > /dev/null 2>&1
cp build/r4/r4_raw.xlsx "siliconomics/docs/reports/r4_bridge/20261118_R4對帳橋接_Siliconomics側_r13-$N.xlsx"
python3 siliconomics/tools/recon.py "siliconomics/docs/reports/r4_bridge/20261118_R4對帳橋接_Siliconomics側_r13-$N.xlsx" siliconomics/data/recon_snapshot_$V.json > build/recon.log 2>&1
echo recon_snapshot_$V.json > siliconomics/data/CURRENT_RECON
cp siliconomics/model/$F build/pass1.xlsx
bld build/pass1.xlsx 2
python3 siliconomics/tools/export_snapshot.py siliconomics/model/$F siliconomics/data/siliconomics_snapshot_$V.json --allow-trial-sgov > build/snap.log 2>&1
echo siliconomics_snapshot_$V.json > siliconomics/data/CURRENT_SNAPSHOT
python3 siliconomics/tools/export_baseline_status.py > build/status.log 2>&1 || echo "status export failed" >> build/status.log
# 診斷（不回寫模型）：收斂後的換算值 x2（協調橋接重算）、供給側證據比例重量測
python3 siliconomics/tools/baseline_bridge.py build/bridge_raw.xlsx > build/bridge.log 2>&1 && soffice $LO --headless --calc --convert-to xlsx:"Calc MS Excel 2007 XML" --outdir build/bridge build/bridge_raw.xlsx > /dev/null 2>&1 || echo bridge_fail >> build/bridge.log
python3 siliconomics/tools/supply_evidence_ratio.py build/ev_raw.xlsx > build/ev.log 2>&1 && soffice $LO --headless --calc --convert-to xlsx:"Calc MS Excel 2007 XML" --outdir build/ev build/ev_raw.xlsx > /dev/null 2>&1 || echo ev_fail >> build/ev.log
echo "DONE $N $(date +%T)"
