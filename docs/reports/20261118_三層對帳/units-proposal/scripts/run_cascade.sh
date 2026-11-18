#!/bin/bash
# 情境複本跑個股 cascade 到收斂（最多 4 輪；某輪全部 SAME 即停），之後抽目標價
N=$1; D=/tmp/pp_$N; R=/home/claude/out/proposal/runs
cd $D; export LO_PROFILE=file:///tmp/lo_cas_$N
for i in 1 2 3 4; do
  bash siliconomics/tools/cascade.sh /tmp/cas_${N}_r$i > $R/cas_${N}_r$i.log 2>&1
  python3 /home/claude/out/verify/extract_tp.py $D $R/tp_${N}_r$i.json > /dev/null 2>&1
  nch=$(grep -cE '^(CHANGED|BUILD_FAIL|CALC_FAIL)' $R/cas_${N}_r$i.log)
  echo "$N round $i changed/fail=$nch $(date +%T)" >> $R/cascade_status.log
  rm -rf /tmp/cas_${N}_r$i/raw
  [ "$nch" = 0 ] && break
done
echo "CASDONE $N" >> $R/cascade_status.log
