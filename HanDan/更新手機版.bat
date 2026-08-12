@echo off
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo 找不到 .venv，請先依 CLAUDE.md 完成環境設定：
    echo   uv venv
    echo   .venv\Scripts\activate
    echo   uv pip install requests openpyxl
    pause
    exit /b 1
)

echo ============================================
echo  步驟 1 / 2　抓取盤後收盤價並更新存摺主檔
echo ============================================
echo.
".venv\Scripts\python.exe" fetch_close.py --html "菡萏咖啡-台股損益存摺.html"
if errorlevel 1 (
    echo.
    echo [中止] 抓價失敗，手機版未重新產生，原有檔案保持不變。
    pause
    exit /b 1
)

echo.
echo ============================================
echo  步驟 2 / 2　產生手機版單檔
echo ============================================
echo.
".venv\Scripts\python.exe" build_mobile.py
if errorlevel 1 (
    echo.
    echo [中止] 手機版產生失敗，請把上方訊息告訴 Claude。
    pause
    exit /b 1
)

echo.
pause
