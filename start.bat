@echo off
REM 進入項目目錄
cd /d "C:\Users\ㄋ\Desktop\後台"

REM 檢查 npm 是否安裝
npm --version
if %errorlevel% neq 0 (
    echo ❌ npm 未安裝或不在 PATH 中
    pause
    exit /b 1
)

echo.
echo ✅ npm 已找到
echo.
echo 正在安裝依賴...
npm install

if %errorlevel% equ 0 (
    echo.
    echo ✅ 依賴安裝成功！
    echo.
    echo 🚀 啟動開發服務器...
    npm start
) else (
    echo ❌ npm install 失敗
    pause
)
