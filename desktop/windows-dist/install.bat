@echo off
REM =============================================================
REM PowerUp Portal -- first-time setup [Windows]
REM
REM Run this once after downloading the folder. It:
REM   1. Verifies Python 3.11+ is installed
REM   2. Creates a virtual environment in .venv\
REM   3. Installs all dependencies
REM
REM Notes:
REM   * Pure ASCII. cmd in default code page mis-parses UTF-8.
REM   * No round parentheses inside echo strings - they break
REM     parenthesized if blocks. Use goto-based flow instead.
REM =============================================================

SETLOCAL
cd /d "%~dp0"

echo.
echo ========================================
echo   PowerUp Portal - First-time setup
echo ========================================
echo.

REM --- 1. Check Python ---
echo [1/3] Checking Python installation...
python --version >NUL 2>&1
if errorlevel 1 goto :no_python

for /f "tokens=2" %%v in ('python --version 2^>^&1') do set PYVER=%%v
echo       Found Python %PYVER%.

REM --- 2. Create venv ---
if exist ".venv\Scripts\python.exe" goto :venv_ready

echo [2/3] Creating virtual environment in .venv\ ...
python -m venv .venv
if errorlevel 1 goto :venv_failed
goto :venv_done

:venv_ready
echo [2/3] Virtual environment already exists, reusing it.

:venv_done

REM --- 3. Install dependencies ---
echo [3/3] Installing dependencies. Takes ~1-2 min on first run...
.venv\Scripts\python.exe -m pip install --upgrade pip --quiet
.venv\Scripts\python.exe -m pip install -r requirements.txt --quiet
if errorlevel 1 goto :pip_failed

echo.
echo ========================================
echo   Setup complete!
echo ========================================
echo.
echo Double-click   run-PowerUp-Portal.bat   to start the app.
echo.
pause
exit /b 0


:no_python
echo.
echo [ERROR] Python is not installed, or not on PATH.
echo.
echo Please install Python 3.11 or later from:
echo     https://www.python.org/downloads/windows/
echo.
echo IMPORTANT: During installation, tick the box
echo     [x] Add Python to PATH
echo.
echo Then re-run this script.
echo.
pause
exit /b 1


:venv_failed
echo.
echo [ERROR] Could not create virtual environment.
pause
exit /b 1


:pip_failed
echo.
echo [ERROR] pip install failed. See messages above.
pause
exit /b 1
