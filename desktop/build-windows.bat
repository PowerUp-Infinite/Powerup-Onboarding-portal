@echo off
REM ─────────────────────────────────────────────────────────────
REM PowerUp Portal — Windows build script.
REM
REM Produces desktop\dist\PowerUp-Portal.exe (single file, ~250 MB).
REM Run from any command prompt; no arguments needed.
REM ─────────────────────────────────────────────────────────────

SETLOCAL
cd /d "%~dp0"

echo.
echo ==== PowerUp Portal — Windows build ====
echo.

REM 1. Ensure Python 3.11+ is available.
python --version >NUL 2>&1
if errorlevel 1 (
    echo [ERROR] Python is not installed or not on PATH.
    echo         Install Python 3.11 from https://www.python.org/downloads/windows/
    echo         and re-run this script.
    exit /b 1
)

REM 2. (Optional) create a venv so we don't pollute the user's global site-packages.
if not exist ".venv\Scripts\python.exe" (
    echo [1/4] Creating virtual environment...
    python -m venv .venv
)

echo [2/4] Installing dependencies...
call .venv\Scripts\activate.bat
python -m pip install --upgrade pip --quiet
python -m pip install -r requirements.txt --quiet
if errorlevel 1 (
    echo [ERROR] pip install failed.
    exit /b 1
)

echo [3/4] Verifying credentials.json is bundled...
if not exist "resources\credentials.json" (
    echo [ERROR] desktop\resources\credentials.json is missing.
    echo         Copy your service account JSON to that path before building.
    exit /b 1
)

echo [4/4] Running PyInstaller...
pyinstaller PowerUp-Portal.spec --clean --noconfirm
if errorlevel 1 (
    echo [ERROR] PyInstaller build failed.
    exit /b 1
)

echo.
echo ==== Build complete ====
echo Output: %cd%\dist\PowerUp-Portal.exe
echo.
echo Next steps:
echo   1. Test by double-clicking dist\PowerUp-Portal.exe
echo   2. Upload dist\PowerUp-Portal.exe to Google Drive for your teammates
echo.
