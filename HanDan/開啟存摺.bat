@echo off
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo Cannot find .venv - please run the setup steps in CLAUDE.md first:
    echo   uv venv
    echo   .venv\Scripts\activate
    echo   uv pip install requests openpyxl
    pause
    exit /b 1
)

".venv\Scripts\python.exe" serve.py
pause
