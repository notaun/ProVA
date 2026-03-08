@echo off
cd /d "%~dp0"

call .venv\Scripts\activate.bat

echo [ProVA] Starting...
python prova_ui.py

if errorlevel 1 (
    echo.
    echo [ProVA exited with an error - see above]
    pause
)