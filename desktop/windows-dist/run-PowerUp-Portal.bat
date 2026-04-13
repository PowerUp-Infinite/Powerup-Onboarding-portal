@echo off
REM ─────────────────────────────────────────────────────────────
REM PowerUp Portal — launcher (Windows)
REM
REM Double-click this file to start the app.
REM If nothing happens, open install.bat first (one-time setup).
REM ─────────────────────────────────────────────────────────────

cd /d "%~dp0"

REM Verify setup has been run
if not exist ".venv\Scripts\pythonw.exe" (
    echo.
    echo [ERROR] It looks like you haven't run install.bat yet.
    echo.
    echo Double-click install.bat first (one-time setup),
    echo then re-run this file.
    echo.
    pause
    exit /b 1
)

REM pythonw.exe runs without a console window. If the app crashes at
REM startup, main.py catches the exception and pops up a Tkinter error
REM dialog — so the user still sees what went wrong.
start "" /B ".venv\Scripts\pythonw.exe" "app\main.py"
exit /b 0
