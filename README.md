# 📄 RAAWA Generator - Automated Multi-Site Access Request Tool

[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://share.streamlit.io)

## Overview

The **RAAWA Generator** automates the generation of RAAWA (Request for Authority to Access Work Area) forms for telecommunications infrastructure projects.

## Quick Start

### Prerequisites
- Python 3.9 or higher

### Installation (No Admin Rights Required)

1. **Clone or download** this repository

2. **Install Python packages**
   ```bash
   pip install --user streamlit pandas openpyxl numpy
Place your Excel files in this folder:

Globe FO Engr Contact_Vendor.xlsx (Site database)

Requisitioner.xlsx (Requisitioner database)

EngrTech.xlsx (Engineer/Technician database)

MIN591_MANUAL RAWA_APPLICATION_June8,2026.xlsx (Template)

Run the application

bash
streamlit run app.py
OR double-click RUN_RAWA.ps1

Open browser to http://localhost:8501

Features
✅ Unlimited site selection with auto-batching (10 sites per RAAWA)

✅ Automatic territory and facility manager conflict resolution

✅ Personnel management (manual or database import)

✅ Auto-format contact numbers (adds leading 0)

✅ Clean ID numbers (removes decimal places)

✅ Professional Excel output with proper styling

File Structure
text
Raawa Generator/
├── app.py                          # Main application
├── requirements.txt                # Python dependencies
├── RUN_RAWA.ps1                    # PowerShell launcher
├── RAWA_Generator_User_Manual.pdf  # User documentation
└── [Your Excel files]              # Keep locally (not on GitHub)
Required Excel Files (Keep Local)
File	Description
Globe FO Engr Contact_Vendor.xlsx	Site database with PLAID, SITE, TERRITORY, etc.
Requisitioner.xlsx	Requisitioner profiles by territory
EngrTech.xlsx	Engineer/Technician database
MIN591_MANUAL RAWA_APPLICATION_June8,2026.xlsx	RAAWA template
Troubleshooting
"Python not found"
Install Python from python.org

Make sure "Add Python to PATH" is checked

"Module not found"
bash
pip install --user streamlit pandas openpyxl numpy
App doesn't open
Verify all Excel files are in the correct folder

Check file names match exactly (case-sensitive)

Developer
John Carlo Rabanes
OLT Rollout Engineer, Nokia Shanghai Bell
📧 rabanes.johncarlo4@gmail.com

Version
v2.2.0 - June 2026

License
Private - For internal use only

text

## File 3: `install.bat` (Create this - alternative to PowerShell)

```batch
@echo off
title RAAWA Generator - Installer
color 0A

echo ============================================================
echo    RAAWA Generator - Package Installer
echo    (No Administrator Rights Required)
echo ============================================================
echo.

cd /d "%~dp0"

:: Check Python
python --version > nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python is not installed!
    echo.
    echo Download Python from: https://www.python.org/downloads/
    echo IMPORTANT: Check "Add Python to PATH" during installation
    echo.
    start https://www.python.org/downloads/
    pause
    exit /b
)

echo [OK] Python found:
python --version
echo.

:: Install packages
echo Installing packages to user folder...
echo.

python -m pip install --user streamlit pandas openpyxl numpy

echo.
echo ============================================================
echo    Installation Complete!
echo ============================================================
echo.
echo Installed packages:
python -c "import streamlit; print('  ✓ streamlit ' + streamlit.__version__)" 2>nul
python -c "import pandas; print('  ✓ pandas ' + pandas.__version__)" 2>nul
python -c "import openpyxl; print('  ✓ openpyxl ' + openpyxl.__version__)" 2>nul
echo.
echo ============================================================
echo.
echo To start the app:
echo   1. Double-click RUN_RAWA.ps1
echo   2. OR run: streamlit run app.py
echo.
pause
