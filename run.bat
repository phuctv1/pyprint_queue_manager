@echo off
chcp 65001 > nul
title PyPrint Queue Manager
echo ===================================================
echo     DANG KHOI DONG PYPRINT QUEUE MANAGER...
echo ===================================================
cd /d "%~dp0"
python main.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Da xay ra loi khi chay ung dung!
    pause
)
