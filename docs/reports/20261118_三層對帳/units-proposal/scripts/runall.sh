#!/bin/bash
# 依序建置 GW 情境；不跑個股的情境建完即刪複本以省空間
cd /tmp
R=/home/claude/out/proposal/runs; S=/home/claude/out/proposal/scripts
run() { # 名 來源 edits pc keep
  bash $S/run_sil.sh $1 $2 $3 full $4 >> $R/runall.log 2>&1 && python3 $S/metrics.py $1 >> $R/runall.log 2>&1 || echo "FAIL $1" >> $R/runall.log
  [ "$5" = keep ] || rm -rf /tmp/pp_$1
}
M=/tmp/pp_src_main; P=/tmp/pp_src_661
run S1 $M $R/edits_S1.json 0 del
run S2 $P $R/edits_S1.json 1 del
run S3a $M $R/edits_S3.json 0 keep
run S3b $P $R/edits_S3.json 1 keep
run S3m_a $M $R/edits_S3m.json 0 del
run S3split_a $M $R/edits_S3split.json 0 del
run S3split_b $P $R/edits_S3split.json 1 del
run S4a $M $R/edits_S4.json 0 del
run S4b $P $R/edits_S4.json 1 keep
run S4h_a $M $R/edits_S4h.json 0 del
run S4h_b $P $R/edits_S4h.json 1 del
run S5L_a $M $R/edits_S5L.json 0 del
run S5L_b $P $R/edits_S5L.json 1 del
run S5H_a $M $R/edits_S5H.json 0 del
run S5H_b $P $R/edits_S5H.json 1 del
echo ALLDONE >> $R/runall.log
