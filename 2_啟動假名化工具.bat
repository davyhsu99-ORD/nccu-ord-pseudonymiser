@echo off
chcp 65001 >nul
cd /d "%~dp0"
title 假名化工具

where pythonw >nul 2>nul
if %errorlevel%==0 (
    start "" pythonw "%~dp0anonymize_gui.py"
    exit /b
)

pyw -3 --version >nul 2>nul
if %errorlevel%==0 (
    start "" pyw -3 "%~dp0anonymize_gui.py"
    exit /b
)

py -3 --version >nul 2>nul
if %errorlevel%==0 (
    py -3 "%~dp0anonymize_gui.py"
    exit /b
)

python --version >nul 2>nul
if %errorlevel%==0 (
    python "%~dp0anonymize_gui.py"
    exit /b
)

echo.
echo ==================================================================
echo   找不到 Python。
echo   請先點兩下「1_第一次執行_安裝環境.bat」完成安裝。
echo ==================================================================
echo.
pause
