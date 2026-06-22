@echo off
title RAAWA Generator
color 0A

cd /d "%~dp0"

echo ============================================================
echo    RAAWA Generator - Starting...
echo ============================================================
echo.

:: Check Python
python --version > nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found!
    echo Please run install.bat first
    pause
    exit /b
)

:: Check packages
python -c "import streamlit" > nul 2>&1
if errorlevel 1 (
    echo Installing missing packages...
    python -m pip install --user streamlit pandas openpyxl
    echo.
)

echo Starting application...
echo.
echo The app will open in your browser
echo Close this window to stop the server
echo.
echo ============================================================
echo.

streamlit run app.py

pause