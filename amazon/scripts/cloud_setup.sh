#!/bin/bash
# Claude Code 雲端環境 Setup script（v3：極簡版）
# v2 在安裝階段卡住超過 45 分鐘：背景安裝的子程序持有輸出管線，執行器等不到結束訊號。
# 因此這裡只做一件快速、不會卡住的事；LibreOffice 與 Chromium 改由 Claude 在工作階段內安裝（CLAUDE.md 第 0 步）。
# 所有輸出寫入檔案、不在背景執行、一律以 0 結束。
# cffi：系統的 cryptography 缺 _cffi_backend 時，import pypdf 會失敗。
# playwright 固定 1.56.0：其 Chromium 為 build 1194（141.0.7390.37），與雲端環境預裝的 /opt/pw-browsers/chromium-1194 相符，
# 不需下載瀏覽器。更新 playwright 前先確認新版的 Chromium build 與環境預裝版本一致。
timeout -k 10 120 pip install -q --break-system-packages openpyxl pypdf playwright==1.56.0 cffi > /tmp/crwv_setup.log 2>&1 || true
exit 0
