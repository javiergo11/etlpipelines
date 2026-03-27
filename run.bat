@echo off
echo ============================================================
echo   EXR Owner Financials Extractor
echo ============================================================
echo.
 
REM Check if Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in your PATH.
    echo Please install Python from https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation.
    pause
    exit /b 1
)
 
REM Install dependencies (only installs if not already present)
echo Installing dependencies...
python -m pip install -r requirements.txt --quiet
echo.
 
REM Run the script
python extract_owner_financials.py
 
echo.
pause
