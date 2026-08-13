@echo off
REM Trendyol QC tool -- convenience wrapper for Windows.
REM
REM Usage (double-click won't work since it needs a file argument -- run
REM from Command Prompt, or drag-and-drop the CSV file onto this .bat file
REM in Windows Explorer):
REM   run_qc.bat trustana_export.csv [--verify-images]
REM
REM The report is saved next to the input file as <name>_QC_Report.xlsx.
REM First run sets up a local Python environment automatically (one-time,
REM ~30 seconds); every run after that is instant.

setlocal enabledelayedexpansion
set DIR=%~dp0

if "%~1"=="" (
  echo Usage: run_qc.bat trustana_export.csv [--verify-images]
  echo The report will be saved next to the input file as ^<name^>_QC_Report.xlsx
  exit /b 1
)

if not exist "%~1" (
  echo Error: can't find file "%~1"
  exit /b 1
)

set INPUT=%~1
set INPUT_DIR=%~dp1
set BASENAME=%~n1
set OUTPUT=%INPUT_DIR%%BASENAME%_QC_Report.xlsx

if not exist "%DIR%.venv" (
  echo First run: setting up Python environment ^(one-time, ~30 seconds^)...
  python -m venv "%DIR%.venv"
  "%DIR%.venv\Scripts\pip" install --quiet --upgrade pip
  "%DIR%.venv\Scripts\pip" install --quiet -r "%DIR%requirements.txt"
)

shift
"%DIR%.venv\Scripts\python" "%DIR%trendyol_qc_tool.py" "%INPUT%" "%OUTPUT%" %1 %2 %3

echo.
echo Done. Report saved to: %OUTPUT%
endlocal
