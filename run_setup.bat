@echo off
title RAAWA Generator Setup
color 0A

:: Keep window open even if error occurs
setlocal enabledelayedexpansion

:MENU
cls
echo ===============================================================================
echo                       RAAWA GENERATOR SETUP
echo ===============================================================================
echo.
echo   This tool will help you install and run the RAAWA Generator
echo.
echo ===============================================================================
echo.
echo   [1] Check Python Installation
echo   [2] Install Required Packages
echo   [3] Run RAAWA Generator
echo   [4] Create Desktop Shortcut
echo   [5] Complete Setup (All Steps)
echo   [6] Exit
echo.
echo ===============================================================================
echo.
set /p choice="Enter your choice (1-6): "

if "%choice%"=="1" goto CHECK_PYTHON
if "%choice%"=="2" goto INSTALL_PACKAGES
if "%choice%"=="3" goto RUN_APP
if "%choice%"=="4" goto CREATE_SHORTCUT
if "%choice%"=="5" goto COMPLETE_SETUP
if "%choice%"=="6" goto EXIT
goto MENU

:CHECK_PYTHON
cls
echo ===============================================================================
echo                        CHECKING PYTHON INSTALLATION
echo ===============================================================================
echo.
python --version > nul 2>&1
if errorlevel 1 (
    echo [FAIL] Python is NOT installed
    echo.
    echo ===============================================================================
    echo                        PYTHON NOT FOUND
    echo ===============================================================================
    echo.
    echo Please install Python manually:
    echo.
    echo   1. Open your web browser
    echo   2. Go to: https://www.python.org/downloads/
    echo   3. Click "Download Python 3.12.3"
    echo   4. Run the installer
    echo   5. IMPORTANT: CHECK "Add Python to PATH"
    echo   6. Click "Install Now"
    echo   7. After installation, RESTART your computer
    echo   8. Run this setup again
    echo.
    echo ===============================================================================
    echo.
    start https://www.python.org/downloads/
) else (
    echo [OK] Python is installed
    python --version
    echo.
    echo Python PATH:
    where python
)
echo.
pause
goto MENU

:INSTALL_PACKAGES
cls
echo ===============================================================================
echo                    INSTALLING REQUIRED PACKAGES
echo ===============================================================================
echo.
echo This will install:
echo   - streamlit (Web app framework)
echo   - pandas (Excel data handling)
echo   - openpyxl (Excel file generation)
echo   - numpy (Data calculations)
echo.
echo ===============================================================================
echo.
set /p confirm="Continue with installation? (Y/N): "
if /i not "%confirm%"=="Y" goto MENU

echo.
echo Upgrading pip...
python -m pip install --upgrade pip
echo.

echo Installing streamlit...
python -m pip install streamlit
echo.

echo Installing pandas...
python -m pip install pandas
echo.

echo Installing openpyxl...
python -m pip install openpyxl
echo.

echo Installing numpy...
python -m pip install numpy
echo.

echo ===============================================================================
echo                    INSTALLATION COMPLETE!
echo ===============================================================================
echo.
pause
goto MENU

:RUN_APP
cls
echo ===============================================================================
echo                        STARTING RAAWA GENERATOR
echo ===============================================================================
echo.
echo Checking if packages are installed...
python -c "import streamlit" > nul 2>&1
if errorlevel 1 (
    echo [WARNING] Packages not found! Installing now...
    python -m pip install streamlit pandas openpyxl
    echo.
)

echo.
echo ===============================================================================
echo                        RAAWA GENERATOR IS STARTING
echo ===============================================================================
echo.
echo The application will open in your web browser
echo.
echo DO NOT CLOSE THIS WINDOW while the app is running
echo.
echo To stop the server, press Ctrl+C
echo.
echo ===============================================================================
echo.

python -m streamlit run app.py

echo.
echo ===============================================================================
echo                    SERVER HAS STOPPED
echo ===============================================================================
echo.
pause
goto MENU

:CREATE_SHORTCUT
cls
echo ===============================================================================
echo                    CREATING DESKTOP SHORTCUT
echo ===============================================================================
echo.

set "DESKTOP=%USERPROFILE%\Desktop"
set "SCRIPT_DIR=%~dp0"
set "SHORTCUT_PATH=%DESKTOP%\RAAWA Generator.bat"

:: Create a simple batch file on desktop
(
echo @echo off
echo cd /d "%SCRIPT_DIR%"
echo python -m streamlit run app.py
echo pause
) > "%SHORTCUT_PATH%"

echo [OK] Desktop shortcut created at: %SHORTCUT_PATH%
echo.
echo You can now double-click "RAAWA Generator" on your desktop
echo.
pause
goto MENU

:COMPLETE_SETUP
cls
echo ===============================================================================
echo                    COMPLETE SETUP - PLEASE WAIT
echo ===============================================================================
echo.

:: Step 1: Check Python
echo [Step 1/3] Checking Python...
python --version > nul 2>&1
if errorlevel 1 (
    echo [FAIL] Python not found!
    echo.
    echo Please install Python first using option 1
    echo.
    pause
    goto MENU
)
echo [OK] Python found
echo.

:: Step 2: Install packages
echo [Step 2/3] Installing packages...
python -m pip install streamlit pandas openpyxl numpy
echo [OK] Packages installed
echo.

:: Step 3: Create shortcut
echo [Step 3/3] Creating desktop shortcut...
set "DESKTOP=%USERPROFILE%\Desktop"
set "SCRIPT_DIR=%~dp0"
set "SHORTCUT_PATH=%DESKTOP%\RAAWA Generator.bat"

(
echo @echo off
echo cd /d "%SCRIPT_DIR%"
echo python -m streamlit run app.py
echo pause
) > "%SHORTCUT_PATH%"
echo [OK] Desktop shortcut created
echo.

echo ===============================================================================
echo                    SETUP COMPLETE!
echo ===============================================================================
echo.
echo You can now:
echo   1. Double-click "RAAWA Generator" on your desktop
echo   2. Or use option 3 from the main menu
echo.
pause
goto MENU

:EXIT
cls
echo ===============================================================================
echo                    THANK YOU FOR USING RAAWA GENERATOR
echo ===============================================================================
echo.
echo To run the app later, double-click "RAAWA Generator" on your desktop
echo.
echo Goodbye!
echo.
timeout /t 3 /nobreak > nul
exit /b