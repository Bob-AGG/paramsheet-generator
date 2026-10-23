@echo off
chcp 65001 >nul
title 参数单批量提取（多线程 + 跳过 PDF）

REM ====== 改这两行 ======
set SRC=D:\
set OUT=C:\Users\Administrator\Desktop\paramsheet_extract
REM =====================

echo ============================================
echo  源目录: %SRC%
echo  输出到: %OUT%
echo  模式  : 只跑 PPT/PPTX（跳过 PDF），8 线程，已提取的自动跳过
echo ============================================
echo.

python extract_batch.py --src "%SRC%" --out "%OUT%" --pptx-only --jobs 8 --skip-existing

echo.
echo 跑完了。输出在: %OUT%
pause
