@echo off
chcp 65001 >nul
title 参数单批量提取

REM ====== 改这两行 ======
set SRC=Z:\参数单\康明斯
set OUT=D:\参数单提取\输出
REM =====================

echo ============================================
echo  源目录: %SRC%
echo  输出到: %OUT%
echo ============================================
echo.

python extract_batch.py --src "%SRC%" --out "%OUT%"

echo.
echo 跑完了。输出在: %OUT%
pause
