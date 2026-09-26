@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
cd /d "%~dp0"
title 假名化工具 - 安裝環境（只需執行一次）

echo.
echo ==================================================================
echo   上傳前假名化工具　環境安裝
echo   這個步驟只需要做一次。裝好之後，以後都只要點「2_啟動假名化工具」。
echo ==================================================================
echo.

:: ---------- 找 Python ----------
set "PY="
py -3 --version >nul 2>nul
if %errorlevel%==0 (
    set "PY=py -3"
    goto :found
)
python --version >nul 2>nul
if %errorlevel%==0 (
    set "PY=python"
    goto :found
)

:: ---------- 沒有 Python ----------
echo [1/3] 這台電腦還沒有安裝 Python。
echo.
winget --version >nul 2>nul
if %errorlevel%==0 (
    echo       可以幫你自動安裝（約 3-5 分鐘，不需要管理員權限）。
    echo.
    choice /c YN /m "      要現在自動安裝 Python 嗎"
    if !errorlevel!==1 (
        echo.
        echo       安裝中，請稍候……
        winget install -e --id Python.Python.3.12 --scope user ^
            --accept-package-agreements --accept-source-agreements
        echo.
        echo ==================================================================
        echo   Python 已安裝完成。
        echo   請「關閉這個視窗」，然後再點一次「1_第一次執行_安裝環境」，
        echo   讓系統重新讀取設定，繼續安裝剩下的套件。
        echo ==================================================================
        echo.
        pause
        exit /b
    )
)

echo.
echo       請手動安裝：
echo         1. 開啟瀏覽器前往　https://www.python.org/downloads/
echo         2. 下載 Python 3.9 以上版本
echo         3. 安裝時務必勾選「Add Python to PATH」
echo         4. 裝完後回來再點一次這個檔案
echo.
pause
exit /b

:found
echo [1/3] 已找到 Python：
%PY% --version
echo.

:: ---------- 安裝套件 ----------
echo [2/3] 安裝必要套件（pandas、openpyxl、xlrd）……
echo.
%PY% -m pip install --upgrade pip >nul 2>nul
%PY% -m pip install pandas openpyxl xlrd
if %errorlevel% neq 0 (
    echo.
    echo   [失敗] 套件安裝沒有成功。
    echo   常見原因是單位網路的防火牆擋住了。請把整個畫面截圖給資訊窗口。
    echo.
    pause
    exit /b
)

echo.
echo       安裝選用套件（讓你可以直接把檔案拖進視窗，裝不起來也不影響使用）……
%PY% -m pip install tkinterdnd2 >nul 2>nul
if %errorlevel%==0 (echo       拖放功能：已啟用) else (echo       拖放功能：未啟用（改用按鈕選檔即可）)

:: ---------- 建立桌面捷徑 ----------
echo.
echo [3/3] 在桌面建立捷徑……
powershell -NoProfile -Command ^
  "$s=(New-Object -COM WScript.Shell).CreateShortcut([Environment]::GetFolderPath('Desktop')+'\假名化工具.lnk');" ^
  "$s.TargetPath='%~dp02_啟動假名化工具.bat';$s.WorkingDirectory='%~dp0';" ^
  "$s.Description='上傳前假名化工具';" ^
  "if (Test-Path '%~dp0假名化工具.ico') { $s.IconLocation='%~dp0假名化工具.ico,0' };" ^
  "$s.Save()" >nul 2>nul
if %errorlevel%==0 (echo       已在桌面建立「假名化工具」捷徑) else (echo       捷徑建立失敗，直接用資料夾裡的 2_啟動假名化工具 即可)

echo.
echo ==================================================================
echo   安裝完成。
echo   請點兩下「2_啟動假名化工具」（或桌面上的捷徑）開始使用。
echo ==================================================================
echo.
pause
