@echo off
REM -------------------------------------------------------------
REM PowerUp Portal - build the Windows distribution folder.
REM
REM Produces a ready-to-zip folder at:
REM   desktop\dist\PowerUp-Portal-Windows\
REM
REM Contents:
REM   install.bat                        (from windows-dist/)
REM   run-PowerUp-Portal.bat             (from windows-dist/)
REM   INSTALL.txt                        (from windows-dist/)
REM   requirements.txt                   (filtered - no pyinstaller)
REM   app\main.py, gui.py, app_config.py, ui_theme.py
REM   app\workers\
REM   app\portal\                        (copy of ../portal/)
REM   app\portal_shims\
REM   app\resources\
REM       credentials.json               (MUST exist in desktop\resources\)
REM       .env                           (copied from ../portal/.env)
REM
REM Zip the result and upload to Google Drive for distribution.
REM -------------------------------------------------------------

SETLOCAL EnableDelayedExpansion
cd /d "%~dp0"

set "OUT=dist\PowerUp-Portal-Windows"
set "APP=%OUT%\app"
set "RES=%APP%\resources"

echo.
echo ========================================
echo   PowerUp Portal - Windows packager
echo ========================================
echo.

REM --- 1. Sanity checks ---
echo [1/6] Sanity checks...
if not exist "..\portal\.env" (
    echo [ERROR] ..\portal\.env is missing. Cannot package.
    exit /b 1
)
if not exist "resources\credentials.json" (
    echo [ERROR] desktop\resources\credentials.json is missing.
    echo         Copy your service account JSON to that path before packaging.
    exit /b 1
)
if not exist "windows-dist\install.bat" (
    echo [ERROR] windows-dist\install.bat is missing. Repo is broken.
    exit /b 1
)

REM --- 2. Clean output folder ---
echo [2/6] Clearing old output folder...
if exist "%OUT%" rmdir /s /q "%OUT%"
mkdir "%OUT%"       >NUL 2>&1
mkdir "%APP%"       >NUL 2>&1
mkdir "%RES%"       >NUL 2>&1

REM --- 3. Copy template files from windows-dist\ ---
echo [3/6] Copying launcher scripts (install.bat, run.bat, INSTALL.txt)...
copy /Y "windows-dist\install.bat"            "%OUT%\"                >NUL
copy /Y "windows-dist\run-PowerUp-Portal.bat" "%OUT%\"                >NUL
copy /Y "windows-dist\INSTALL.txt"            "%OUT%\"                >NUL

REM --- 4. Filter requirements (drop pyinstaller - users don't build) ---
REM Use findstr directly. Earlier `echo %%L` approach broke spectacularly
REM because cmd interprets `>=` in `echo customtkinter>=5.2.2` as
REM redirection, creating files literally named `=5.2.2`.
echo [4/6] Generating filtered requirements.txt...
findstr /V /I /B /C:"pyinstaller" requirements.txt > "%OUT%\requirements.txt"

REM --- 5. Copy the Python code ---
echo [5/6] Copying Python code into app\ ...

REM Root desktop files
copy /Y "main.py"       "%APP%\"    >NUL
copy /Y "gui.py"        "%APP%\"    >NUL
copy /Y "app_config.py" "%APP%\"    >NUL
copy /Y "ui_theme.py"   "%APP%\"    >NUL

REM workers/
xcopy /E /I /Y /Q "workers"      "%APP%\workers"      >NUL

REM portal_shims/
xcopy /E /I /Y /Q "portal_shims" "%APP%\portal_shims" >NUL

REM portal/ - copy ONLY the specific files desktop needs. Whitelist rather
REM than blacklist, so stray dev scripts / Streamlit tabs / cached assets
REM in portal/ never sneak into the distribution.
mkdir "%APP%\portal" >NUL 2>&1
copy /Y "..\portal\m2_engine.py" "%APP%\portal\" >NUL
copy /Y "..\portal\m3_engine.py" "%APP%\portal\" >NUL
copy /Y "..\portal\sheets.py"    "%APP%\portal\" >NUL
copy /Y "..\portal\config.py"    "%APP%\portal\" >NUL

REM resources/ - use powershell for the .env copy because cmd's `copy`
REM mis-parses a destination path ending in `\.env` as a glob.
copy /Y "resources\credentials.json" "%RES%\" >NUL
powershell -NoProfile -Command "Copy-Item -LiteralPath '..\portal\.env' -Destination '%RES%\.env' -Force"
if not exist "%RES%\.env" (
    echo [ERROR] Could not copy ..\portal\.env into %RES%\.env
    exit /b 1
)

REM Strip any __pycache__ dirs that snuck in
for /r "%APP%" %%d in (__pycache__) do (
    if exist "%%d" rmdir /s /q "%%d" >NUL 2>&1
)

REM --- 6. Report ---
echo [6/6] Done.
echo.
echo ========================================
echo   Package built successfully
echo ========================================
echo.
echo Output folder:  %~dp0%OUT%
echo.
echo Next:
echo   1. Right-click the folder - "Send to" - "Compressed (zipped) folder"
echo   2. Upload the .zip to your shared Google Drive
echo   3. Teammates download, extract, open INSTALL.txt, follow steps
echo.
exit /b 0
