@echo off
title RAAWA Generator
color 0A
cd /d "%~dp0"

echo ===============================================================================
echo                         RAAWA GENERATOR
echo ===============================================================================
echo.
echo Starting application...
echo.
echo The app will open in your browser
echo Close this window to stop the server
echo.
echo ===============================================================================
echo.

python -m streamlit run app.py

echo.
echo Server stopped.
echo.
pause