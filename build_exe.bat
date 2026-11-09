@echo off
title eFBR CPR Downloader - Build Script
echo ============================================================
echo   eFBR CPR Receipt Downloader - EXE Build Script
echo ============================================================
echo.

:: ── Step 0: Check Python ─────────────────────────────────────────
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH.
    echo Please install Python 3.10+ from https://python.org
    echo Make sure to check "Add Python to PATH" during install.
    pause
    exit /b 1
)

:: ── Step 1: Install dependencies ─────────────────────────────────
echo [1/3] Installing Python dependencies...
python -m pip install --upgrade pip --quiet
python -m pip install -r requirements-build.txt --quiet
echo       Done.

:: ── Step 2: Clean previous build ─────────────────────────────────
echo.
echo [2/3] Cleaning previous build artifacts...
if exist dist rmdir /s /q dist
if exist build rmdir /s /q build
if exist eFBR_CPR_Downloader.spec del /q eFBR_CPR_Downloader.spec
echo       Done.

:: ── Step 3: Build EXE ────────────────────────────────────────────
echo.
echo [3/3] Building standalone EXE (this may take 2-3 minutes)...
python -m PyInstaller ^
    --onefile ^
    --noconsole ^
    --name "eFBR_CPR_Downloader" ^
    --add-data "constants.py;." ^
    --add-data "excel_handler.py;." ^
    --add-data "browser_engine.py;." ^
    --add-data "report_writer.py;." ^
    --add-data "gui.py;." ^
    --add-data "themes.py;." ^
    --hidden-import "customtkinter" ^
    --hidden-import "playwright" ^
    --hidden-import "pandas" ^
    --hidden-import "openpyxl" ^
    --hidden-import "xlrd" ^
    --hidden-import "bootstrap" ^
    --add-data "bootstrap.py;." ^
    --collect-all "customtkinter" ^
    --collect-all "playwright" ^
    main.py

if errorlevel 1 (
    echo.
    echo ERROR: Build failed. See output above for details.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo   BUILD COMPLETE!
echo ============================================================
echo.
echo   Your EXE is ready at:
echo     dist\eFBR_CPR_Downloader.exe
echo.
echo   NOTES:
echo   - The EXE is small: the Chromium browser is NOT bundled.
echo   - On first launch it downloads Chromium by itself (~150 MB,
echo     internet required, one-time only), showing a progress window.
echo   - No Python, Chrome or Edge needed on the target PC.
echo.
pause
