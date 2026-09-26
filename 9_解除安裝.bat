@echo off
chcp 65001 >nul
cd /d "%~dp0"
title 假名化工具 - 解除安裝

:: 只負責找 Python 並把工作交給 uninstall.py。
:: 所有判斷與刪除邏輯都寫在 Python 裡，批次檔不碰路徑與使用者輸入。
:: 一定要用有主控台的 python／py，不能用 pythonw（解除安裝需要問問題）。

if not exist "%~dp0uninstall.py" goto :nopy_file

py -3 --version >nul 2>nul
if %errorlevel%==0 (
    py -3 -X utf8 "%~dp0uninstall.py"
    goto :eof
)

python --version >nul 2>nul
if %errorlevel%==0 (
    python -X utf8 "%~dp0uninstall.py"
    goto :eof
)

echo.
echo ==================================================================
echo   找不到 Python，所以沒有辦法執行解除安裝程式。
echo.
echo   這不影響你移除這個工具。請自己做這三件事：
echo.
echo     1. 先備份：把這個資料夾裡的「專案」整個複製到別的地方。
echo        裡面的 _private\mapping.csv 是代碼與真實姓名的唯一對照表，
echo        刪掉就永遠還原不回姓名，明年的年度作業也會接不上。
echo.
echo     2. 刪掉桌面上的「假名化工具」捷徑。
echo.
echo     3. 確認備份沒問題之後，再把整個工具資料夾刪掉。
echo.
echo   Python 本體與 pandas 等套件請不要隨意移除，
echo   這台電腦的其他程式可能也在用。
echo ==================================================================
echo.
pause
goto :eof

:nopy_file
echo.
echo ==================================================================
echo   找不到 uninstall.py。
echo   請確認這個檔案和 uninstall.py 在同一個資料夾裡。
echo ==================================================================
echo.
pause
