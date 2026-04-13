@echo off
REM =============================================================
REM PowerUp Portal -- first-time setup [Windows]
REM
REM  1. If Python 3.11 is missing, silently installs it from the
REM     bundled python-3.11.9-amd64.exe (no separate download needed).
REM  2. Creates a virtual environment in .venv\
REM  3. Installs all dependencies into the venv
REM
REM  Pure ASCII, CRLF, no parens inside echo strings -- cmd-safe.
REM =============================================================

SETLOCAL
cd /d "%~dp0"

echo.
echo ========================================
echo   PowerUp Portal - First-time setup
echo ========================================
echo.

REM --- 1. Check Python ---
echo [1/4] Checking Python installation...
python --version >NUL 2>&1
if errorlevel 1 goto :install_python
goto :have_python

:install_python
echo       Python not found. Installing from bundled installer...
echo       This is silent and takes about 1 minute.
if not exist "python-3.11.9-amd64.exe" goto :installer_missing

REM Per-user install -- no admin needed. PrependPath puts python on PATH.
"python-3.11.9-amd64.exe" /quiet InstallAllUsers=0 PrependPath=1 Include_test=0
if errorlevel 1 goto :py_install_failed

REM Python's installer adds entries to PATH but the current cmd session
REM doesn't see them until we refresh manually. Add the standard
REM per-user install location to PATH for THIS session.
set "PATH=%LOCALAPPDATA%\Programs\Python\Python311;%LOCALAPPDATA%\Programs\Python\Python311\Scripts;%PATH%"

REM Verify the install actually worked
python --version >NUL 2>&1
if errorlevel 1 goto :py_install_failed
echo       Python installed successfully.

:have_python
for /f "tokens=2" %%v in ('python --version 2^>^&1') do set PYVER=%%v
echo       Python version: %PYVER%

REM --- 2. Create venv ---
if exist ".venv\Scripts\python.exe" goto :venv_ready

echo [2/4] Creating virtual environment in .venv\ ...
python -m venv .venv
if errorlevel 1 goto :venv_failed
goto :venv_done

:venv_ready
echo [2/4] Virtual environment already exists, reusing it.

:venv_done

REM --- 3. Upgrade pip ---
echo [3/4] Upgrading pip...
.venv\Scripts\python.exe -m pip install --upgrade pip --quiet

REM --- 4. Install dependencies ---
echo [4/4] Installing dependencies. Takes ~1-2 min on first run...
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


:installer_missing
echo.
echo [ERROR] python-3.11.9-amd64.exe is missing from this folder.
echo         Re-download the PowerUp-Portal-Windows.zip and try again.
echo.
pause
exit /b 1


:py_install_failed
echo.
echo [ERROR] Could not install Python automatically.
echo.
echo Please install manually:
echo   1. Double-click python-3.11.9-amd64.exe
echo   2. Tick the box "Add python.exe to PATH" at the bottom
echo   3. Click "Install Now"
echo   4. After it finishes, re-run this install.bat
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
