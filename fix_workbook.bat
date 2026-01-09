@echo off
REM Tableau Workbook Fixer - Windows Batch Script
REM Usage: Drag and drop your .twbx file onto this batch file

setlocal enabledelayedexpansion

if "%~1"=="" (
    echo.
    echo Tableau Workbook Error Fixer
    echo ============================
    echo.
    echo Usage: Drag and drop your .twbx file onto this batch file
    echo   OR
    echo Run: fix_workbook.bat "path\to\your\workbook.twbx"
    echo.
    pause
    exit /b 1
)

set "WORKBOOK_PATH=%~1"

echo.
echo Tableau Workbook Error Fixer
echo ============================
echo.
echo Workbook: %WORKBOOK_PATH%
echo.

REM Check if file exists
if not exist "%WORKBOOK_PATH%" (
    echo ERROR: File not found: %WORKBOOK_PATH%
    echo.
    pause
    exit /b 1
)

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo.
    echo Please install Python from: https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation
    echo.
    pause
    exit /b 1
)

echo Running fix script...
echo.

REM Get the directory where this batch file is located
set "SCRIPT_DIR=%~dp0"

REM Run the Python script
python "%SCRIPT_DIR%fix_tableau_workbook.py" "%WORKBOOK_PATH%"

set EXIT_CODE=%ERRORLEVEL%

echo.
if %EXIT_CODE%==0 (
    echo ============================
    echo SUCCESS! Your workbook has been fixed.
    echo ============================
) else (
    echo ============================
    echo ERROR: Fix failed. See above for details.
    echo ============================
)

echo.
pause
exit /b %EXIT_CODE%
