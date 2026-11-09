@echo off
title eFBR CPR Downloader - Setup
cd /d "%~dp0"
echo ============================================================
echo   eFBR CPR Receipt Downloader - Setup
echo   (Optional: run.bat does this automatically on first start)
echo ============================================================
echo.

call :find_python
if not defined PY goto :no_python

echo Using: %PY%
echo Downloading Python packages and Chromium browser...
%PY% main.py --setup-only
if errorlevel 1 (
    echo.
    echo Setup failed. Check your internet connection and try again.
    pause
    exit /b 1
)
echo.
echo ============================================================
echo   Setup complete! Double-click "run.bat" to start the app.
echo ============================================================
pause
exit /b 0

:no_python
echo Python was not found. Run "run.bat" - it can install Python for you.
pause
exit /b 1

:find_python
set "PY="
py -3 --version >nul 2>&1 && set "PY=py -3"
if defined PY exit /b 0
python --version >nul 2>&1 && set "PY=python"
exit /b 0
