#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Fixes XML schema validation errors in Tableau workbook files.

.DESCRIPTION
    This PowerShell script runs the Python workbook fixer to resolve common
    XML errors in Tableau .twbx files, including invalid UUIDs, undeclared
    elements, and content model violations.

.PARAMETER WorkbookPath
    Path to the Tableau workbook (.twbx) file to fix.

.EXAMPLE
    .\fix_workbook.ps1 "C:\Users\yoga sundaram\Downloads\Covid Dashboard - Enhanced V2.twbx"

.EXAMPLE
    .\fix_workbook.ps1 -WorkbookPath ".\MyWorkbook.twbx"
#>

param(
    [Parameter(Mandatory=$true, Position=0, HelpMessage="Path to the Tableau workbook file")]
    [string]$WorkbookPath
)

# Script configuration
$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$PythonScript = Join-Path $ScriptDir "fix_tableau_workbook.py"

# Display header
Write-Host ""
Write-Host "Tableau Workbook Error Fixer" -ForegroundColor Cyan
Write-Host "============================" -ForegroundColor Cyan
Write-Host ""

# Validate workbook file exists
if (-not (Test-Path $WorkbookPath)) {
    Write-Host "ERROR: File not found: $WorkbookPath" -ForegroundColor Red
    Write-Host ""
    exit 1
}

# Get absolute path
$WorkbookPath = Resolve-Path $WorkbookPath

Write-Host "Workbook: " -NoNewline
Write-Host $WorkbookPath -ForegroundColor Yellow
Write-Host ""

# Check if Python is installed
try {
    $pythonVersion = & python --version 2>&1
    Write-Host "Python: " -NoNewline
    Write-Host $pythonVersion -ForegroundColor Green
} catch {
    Write-Host "ERROR: Python is not installed or not in PATH" -ForegroundColor Red
    Write-Host ""
    Write-Host "Please install Python from: https://www.python.org/downloads/" -ForegroundColor Yellow
    Write-Host "Make sure to check 'Add Python to PATH' during installation" -ForegroundColor Yellow
    Write-Host ""
    exit 1
}

# Check if fix script exists
if (-not (Test-Path $PythonScript)) {
    Write-Host "ERROR: Python script not found: $PythonScript" -ForegroundColor Red
    Write-Host ""
    exit 1
}

Write-Host ""
Write-Host "Running fix script..." -ForegroundColor Cyan
Write-Host ""

# Run the Python script
try {
    & python $PythonScript $WorkbookPath

    if ($LASTEXITCODE -eq 0) {
        Write-Host ""
        Write-Host "============================" -ForegroundColor Green
        Write-Host "SUCCESS! Workbook has been fixed." -ForegroundColor Green
        Write-Host "============================" -ForegroundColor Green
        Write-Host ""

        # Show the output file location
        $outputFile = [System.IO.Path]::GetFileNameWithoutExtension($WorkbookPath) + "_fixed.twbx"
        $outputPath = Join-Path (Split-Path $WorkbookPath) $outputFile

        if (Test-Path $outputPath) {
            Write-Host "Fixed workbook location:" -ForegroundColor Cyan
            Write-Host $outputPath -ForegroundColor Yellow
            Write-Host ""

            # Ask if user wants to open the folder
            $response = Read-Host "Open containing folder? (Y/N)"
            if ($response -eq 'Y' -or $response -eq 'y') {
                Start-Process explorer.exe "/select,`"$outputPath`""
            }
        }

        exit 0
    } else {
        Write-Host ""
        Write-Host "============================" -ForegroundColor Red
        Write-Host "ERROR: Fix failed." -ForegroundColor Red
        Write-Host "============================" -ForegroundColor Red
        Write-Host ""
        exit 1
    }
} catch {
    Write-Host ""
    Write-Host "ERROR: $_" -ForegroundColor Red
    Write-Host ""
    exit 1
}
