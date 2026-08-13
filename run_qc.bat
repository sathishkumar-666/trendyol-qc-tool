@echo off
REM Trustana AI Content Verifier for MP -- convenience wrapper for Windows.
REM
REM Usage (double-click won't work since it needs a file argument -- run
REM from Command Prompt, or drag-and-drop the CSV file onto this .bat file
REM in Windows Explorer, which defaults to the Trendyol marketplace):
REM   run_qc.bat trustana_export.csv [trendyol|noon] [--verify-images]
REM                                  [--export-template OUTPUT_XLSX]
REM                                  [--template-file CURRENT_TEMPLATE_XLSX] [--allow-soft-issues]
REM
REM Marketplace defaults to "trendyol" if omitted. The report is saved next
REM to the input file as <name>_<marketplace>_QC_Report.xlsx. First run sets
REM up a local Python environment automatically (one-time, ~30 seconds);
REM every run after that is instant.
REM
REM --export-template OUTPUT_XLSX     Also fill the marketplace's own upload
REM                                    template with every QC-clean product.
REM --template-file CURRENT_TEMPLATE  The upload template to fill -- download
REM                                    the current one from the marketplace's
REM                                    seller center first (Trendyol's varies
REM                                    by which categories you requested it
REM                                    for). Omit to fall back to the bundled
REM                                    copy (may be stale).
REM --allow-soft-issues                Relax "QC-clean" for --export-template
REM                                    from zero issues to zero rejection-risk
REM                                    (HARD) issues. See README's "Stage 2".
REM
REM Example:
REM   run_qc.bat Products_export.csv noon --export-template noon_upload.xlsx --template-file current_noon_template.xlsx --allow-soft-issues

setlocal enabledelayedexpansion
set DIR=%~dp0

if "%~1"=="" (
  echo Usage: run_qc.bat trustana_export.csv [trendyol^|noon] [--verify-images]
  echo                                       [--export-template OUTPUT_XLSX] [--allow-soft-issues]
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
"%DIR%.venv\Scripts\python" "%DIR%qc_engine.py" "%INPUT%" "%OUTPUT%" --marketplace %MARKETPLACE% %1 %2 %3 %4 %5 %6 %7

echo.
echo Done. Report saved to: %OUTPUT%
endlocal
