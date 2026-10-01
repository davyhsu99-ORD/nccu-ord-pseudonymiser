@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0.."
title 還原 GitHub 設定檔

echo.
echo ==================================================================
echo   還原三個以點開頭的 GitHub 設定檔
echo.
echo   Windows 檔案總管預設會隱藏以點開頭的檔案，壓縮與解壓縮時

echo   也容易漏掉，所以這三個檔案先改成一般檔名存放。

echo   這支程式會把它們放回正確的位置與檔名。
echo ==================================================================
echo.

if not exist "_GitHub設定檔\gitignore.txt" (

    echo [失敗] 找不到 _GitHub設定檔 資料夾，請確認這支程式放在原本的位置。
    pause
    exit /b
)

copy /y "_GitHub設定檔\gitignore.txt" ".gitignore" >nul

copy /y "_GitHub設定檔\gitattributes.txt" ".gitattributes" >nul
if not exist ".github\workflows" mkdir ".github\workflows"
copy /y "_GitHub設定檔\demo-check.yml" ".github\workflows\demo-check.yml" >nul

echo   已還原：

if exist ".gitignore"                      (echo     .gitignore)                      else (echo     [缺] .gitignore)

if exist ".gitattributes"                  (echo     .gitattributes)                  else (echo     [缺] .gitattributes)

if exist ".github\workflows\demo-check.yml" (echo     .github\workflows\demo-check.yml) else (echo     [缺] .github\workflows\demo-check.yml)

echo.
echo ==================================================================
echo   完成。接下來：

echo     1. 這個 _GitHub設定檔 資料夾可以整個刪掉（內容已複製出去）

echo     2. 回到上一層，用 GitHub Desktop 或瀏覽器上傳
echo ==================================================================
echo.
pause
