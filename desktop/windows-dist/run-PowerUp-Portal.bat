@echo off
REM =============================================================
REM PowerUp Portal -- launcher [Windows]
REM
REM Double-click this file to start the app.
REM If nothing happens, run install.bat first one-time setup.
REM
REM Notes:
REM   * Pure ASCII only. cmd in default code page mis-parses UTF-8.
REM   * No round parentheses inside echo strings. Inside an
REM     if-block the closing ) is consumed by cmd's parser even
REM     when it lives in an echo argument, breaking the block.
REM   * goto-based flow used instead of multi-line "if (...) (...)"
REM     for the same reason - safer across shells.
REM =============================================================

cd /d "%~dp0"

if not exist ".venv\Scripts\pythonw.exe" goto :no_setup

REM pythonw.exe runs without a console window. If the app crashes
REM at startup, main.py shows a Tk error dialog so the user still
REM sees what went wrong.
start "" /B ".venv\Scripts\pythonw.exe" "app\main.py"
exit /b 0

:no_setup
echo.
echo [ERROR] It looks like you haven't run install.bat yet.
echo.
echo Double-click install.bat first - this is the one-time setup.
echo Then come back and double-click this file again.
echo.
pause
exit /b 1
