@echo off
setlocal
cd /d "%~dp0"
python -c "import openpyxl, numpy, scipy, bs4" >nul 2>nul || goto :deps
if not exist archive mkdir archive
copy /y "Grayson_Swim_Site_Master.xlsx" "archive\Master_before_rebuild_%RANDOM%.xlsx" >nul
python "scripts\import_record_changes.py"
if errorlevel 1 goto :error
python "scripts\update_master.py" --no-import
if errorlevel 1 goto :error
python "scripts\export_current_records.py"
if errorlevel 1 goto :error
python "scripts\build_portable.py"
if errorlevel 1 goto :error
echo Done - view index.html in the repository root.
pause
exit /b 0
:deps
echo Python or required libraries missing: openpyxl numpy scipy beautifulsoup4.
pause
exit /b 1
:error
echo Rebuild failed; check errors above and archive backup.
pause
exit /b 1
