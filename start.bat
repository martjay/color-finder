@echo off
chcp 65001 >nul

if not exist venv (
    echo [错误] 请先运行 install.bat
    pause
    exit /b 1
)

call venv\Scripts\activate.bat
python app.py
