@echo off
setlocal

cd /d "%~dp0"

streamlit run streamlit_app.py

if errorlevel 1 (
    echo.
    echo 啟動失敗，請先執行：pip install -r requirements.txt
    pause
    exit /b 1
)
