# RAAWA Generator Launcher
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "    RAAWA Generator" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Get the script directory
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $scriptDir

Write-Host "Current Directory: $scriptDir" -ForegroundColor Yellow
Write-Host ""

# Check if Python is installed
try {
    $pythonVersion = python --version 2>&1
    Write-Host "Python found: $pythonVersion" -ForegroundColor Green
} catch {
    Write-Host "ERROR: Python not found!" -ForegroundColor Red
    Write-Host ""
    Write-Host "Please install Python from: https://www.python.org/downloads/" -ForegroundColor Yellow
    Write-Host "Make sure to check 'Add Python to PATH' during installation" -ForegroundColor Yellow
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit
}

Write-Host ""

# Check if streamlit is installed
$streamlitCheck = python -c "import streamlit" 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Host "Installing required packages..." -ForegroundColor Yellow
    python -m pip install streamlit pandas openpyxl
    Write-Host ""
}

# Launch the app
Write-Host "Starting RAAWA Generator..." -ForegroundColor Green
Write-Host ""
Write-Host "The app will open in your browser" -ForegroundColor Cyan
Write-Host "Press Ctrl+C to stop the server" -ForegroundColor Cyan
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

python -m streamlit run app.py

Write-Host ""
Write-Host "Server stopped" -ForegroundColor Red
Read-Host "Press Enter to exit"