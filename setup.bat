@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo.
echo  ╔════════════════════════════════╗
echo  ║        ProVA — First-Time Setup          ║
echo  ╚════════════════════════════════╝
echo.

:: ── 1. Check Python is installed ──────────────────────────────────
echo [1/5] Checking Python...

python --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo  [ERROR] Python was not found.
    echo.
    echo  Please install Python 3.10 or 3.11 from:
    echo  https://www.python.org/downloads/
    echo.
    echo  IMPORTANT: tick "Add Python to PATH" during install.
    echo.
    pause
    exit /b 1
)

:: Check version is 3.9+
for /f "tokens=2 delims= " %%v in ('python --version 2^>^&1') do set "PY_VER=%%v"
for /f "tokens=1,2 delims=." %%a in ("!PY_VER!") do (
    set "PY_MAJOR=%%a"
    set "PY_MINOR=%%b"
)
if !PY_MAJOR! LSS 3 goto :bad_version
if !PY_MAJOR! EQU 3 if !PY_MINOR! LSS 9 goto :bad_version
echo  [OK] Found Python !PY_VER!
goto :python_ok

:bad_version
echo.
echo  [ERROR] Python !PY_VER! is too old. ProVA needs Python 3.9 or newer.
echo  Download from: https://www.python.org/downloads/
echo.
pause
exit /b 1

:python_ok

:: ── 2. Create virtual environment ─────────────────────────────────
echo.
echo [2/5] Creating virtual environment...

if exist ".venv\Scripts\python.exe" (
    echo  [OK] .venv already exists — skipping creation.
) else (
    python -m venv .venv
    if errorlevel 1 (
        echo  [ERROR] Failed to create virtual environment.
        pause
        exit /b 1
    )
    echo  [OK] Virtual environment created.
)

:: ── 3. Install dependencies ────────────────────────────────────────
echo.
echo [3/5] Installing dependencies ^(this may take a few minutes^)...
echo.

.venv\Scripts\python.exe -m pip install --upgrade pip --quiet

:: PyAudio often fails with plain pip on Windows — try pipwin fallback
.venv\Scripts\pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo  [WARN] Standard install had errors. Trying PyAudio fallback...
    .venv\Scripts\pip install pipwin
    .venv\Scripts\python.exe -m pipwin install pyaudio
    :: Retry requirements without PyAudio causing the failure
    .venv\Scripts\pip install -r requirements.txt --ignore-installed pyaudio
)

echo.
echo  [OK] Dependencies installed.

:: ── 4. pywin32 post-install step ──────────────────────────────────
echo.
echo [4/5] Running pywin32 post-install...
.venv\Scripts\python.exe -m pywin32_postinstall -install >nul 2>&1
echo  [OK] Done.

:: ── 5. Create desktop shortcut ────────────────────────────────────
echo.
echo [5/5] Creating desktop shortcut...

set "PROVA_DIR=%~dp0"
:: Remove trailing backslash
if "%PROVA_DIR:~-1%"=="\" set "PROVA_DIR=%PROVA_DIR:~0,-1%"

set "TARGET=%PROVA_DIR%\ProVA.bat"
set "SHORTCUT=%USERPROFILE%\Desktop\ProVA.lnk"

:: Include icon if it exists, otherwise skip
set "ICON_ARG="
if exist "%PROVA_DIR%\prova_icon.ico" (
    set "ICON_ARG=$Shortcut.IconLocation = '%PROVA_DIR%\prova_icon.ico';"
)

powershell -NoProfile -Command "$WshShell = New-Object -ComObject WScript.Shell; $Shortcut = $WshShell.CreateShortcut('%SHORTCUT%'); $Shortcut.TargetPath = '%TARGET%'; $Shortcut.WorkingDirectory = '%PROVA_DIR%'; %ICON_ARG% $Shortcut.Save()"

if exist "%SHORTCUT%" (
    echo  [OK] Desktop shortcut created.
) else (
    echo  [WARN] Could not create desktop shortcut. Run ProVA.bat directly.
)

:: ── Setup complete ─────────────────────────────────────────────────
echo.
echo  ╔═════════════════════════════════╗
echo  ║           Setup complete!                ║
echo  ║                                          ║
echo  ║  Before starting ProVA:                  ║
echo  ║  · Open .env and add your Gmail address  ║
echo  ║    and App Password.                     ║
echo  ║  · (Email won't work without this —      ║
echo  ║    everything else will.)                ║
echo  ║                                          ║
echo  ║  Then double-click ProVA on your Desktop ║
echo  ║  or run ProVA.bat directly.              ║
echo  ╚═════════════════════════════════╝
echo.
pause