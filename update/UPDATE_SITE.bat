@echo off
setlocal EnableExtensions
cd /d "%~dp0"
echo ============================================================
echo Grayson swim site - full data refresh
 echo ============================================================
echo.
echo This will: apply reviewed NZ record changes, then acquire
 echo the eight events for active age-group stars, then Grayson's complete history,
echo recalculate the private Excel workbook, and rebuild portable HTML.
echo.
echo IMPORTANT: Do not commit the private workbook to a PUBLIC GitHub repo.
echo.
where python >nul 2>nul || goto :nopython
python -c "import openpyxl, numpy, scipy, bs4" >nul 2>nul || goto :nodeps
if not exist "archive" mkdir "archive"
set "STAMP=%RANDOM%_%RANDOM%"
copy /y "Grayson_Swim_Site_Master.xlsx" "archive\Master_before_update_%STAMP%.xlsx" >nul
if errorlevel 1 goto :error
echo.
echo Applying approved manual record changes, if any...
python "scripts\import_record_changes.py"
if errorlevel 1 goto :error
echo.
echo STEP 1: selected event histories for current age-group stars only.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\swimcloud_update.ps1"
if errorlevel 1 goto :error
if not exist "input\latest_selected_zip.txt" goto :error
set /p SEL=<"input\latest_selected_zip.txt"
if not exist "%SEL%" goto :error
echo.
echo STEP 2: all available events for Grayson's two SwimCloud profiles.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\swimcloud_grayson_all_events.ps1"
if errorlevel 1 goto :error
if not exist "input\latest_grayson_zip.txt" goto :error
set /p GR=<"input\latest_grayson_zip.txt"
if not exist "%GR%" goto :error
python "scripts\update_master.py" "%SEL%" "%GR%"
if errorlevel 1 goto :error
echo.
echo STEP 3: create offline, self-contained HTML files.
python "scripts\export_current_records.py"
if errorlevel 1 goto :error
python "scripts\build_portable.py"
if errorlevel 1 goto :error
echo.
echo COMPLETE: check the five HTML files in the repository ROOT.
echo Check the record tables, current Top 20, dates, heatmaps, charts,
echo media, and QA_Checks sheet before committing public files.
pause
exit /b 0
:nopython
echo Python is not installed or not on PATH. The browser-based acquisition
 echo can still be run using RUN_ACQUISITION_ONLY.bat without Python.
pause
exit /b 1
:nodeps
echo The rebuild needs Python packages: openpyxl, numpy, scipy, beautifulsoup4.
echo Install once in your Python environment:
echo     py -m pip install openpyxl numpy scipy beautifulsoup4
 echo The acquisition scripts themselves have no Python dependencies.
pause
exit /b 1
:error
echo.
echo UPDATE FAILED. Check the messages above. A workbook backup was saved
 echo in update\archive before modifying any data. Do not publish an unreviewed build.
pause
exit /b 1
