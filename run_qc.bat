@echo off
REM Trustana AI Content Verifier for MP -- convenience wrapper for Windows.
REM
REM Usage (double-click won't work since it needs a file argument -- run
REM from Command Prompt, or drag-and-drop the CSV file onto this .bat file
REM in Windows Explorer, which defaults to the Trendyol marketplace):
REM   run_qc.bat trustana_export.csv [trendyol|noon] [--verify-images]
REM
REM Marketplace defaults to "trendyol" if omitted. The report is saved next
REM to the input file as <name>_<marketplace>_QC_Report.xlsx. First run sets
REM up a local Python environment automatically (one-time, ~30 seconds);
REM every run after that is instant.

setlocal enabledelayedexpansion
set DIR=%~dp0

if "%~1"=="" (
  echo Usage: run_qc.bat trustana_export.csv [trendyol^|noon] [--verify-images]
  echo Marketplace defaults to "trendyol" if omitted.
  echo The report will be saved next to the input file as ^<name^>_^<marketplace^>_QC_Report.xlsx
  exit /b 1
)

if not exist "%~1" (
  echo Error: can't find file "%~1"
  exit /b 1
)

set INPUT=%~1
set INPUT_DIR=%~dp1
set BASENAME=%~n1
shift

set MARKETPLACE=trendyol
if /I "%~1"=="trendyol" (
  set MARKETPLACE=trendyol
  shift
) else if /I "%~1"=="noon" (
  set MARKETPLACE=noon
  shift
)

set OUTPUT=%INPUT_DIR%%BASENAME%_%MARKETPLACE%_QC_Report.xlsx

if not exist "%DIR%.venv" (
  echo First run: setting up Python environment ^(one-time, ~30 seconds^)...
  python -m venv "%DIR%.venv"
  "%DIR%.venv\Scripts\pip" install --quiet --upgrade pip
  "%DIR%.venv\Scripts\pip" install --quiet -r "%DIR%requirements.txt"
)

echo Marketplace: %MARKETPLACE%
"%DIR%.venv\Scripts\python" "%DIR%qc_engine.py" "%INPUT%" "%OUTPUT%" --marketplace %MARKETPLACE% %1 %2 %3

echo.
echo Done. Report saved to: %OUTPUT%
endlocal
