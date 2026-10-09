@echo off
setlocal
cd /d "%~dp0"
echo SwimCloud acquisition only - requires no Python or browser extensions.
echo Be ready to clear Cloudflare and press ENTER in the command window.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\swimcloud_update.ps1"
if errorlevel 1 goto :error
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\swimcloud_grayson_all_events.ps1"
if errorlevel 1 goto :error
echo Acquisition finished. ZIP outputs are in update\scripts.
echo The master workbook and public pages have NOT yet been refreshed.
pause
exit /b 0
:error
echo Acquisition failed. Review script logs and any partial output ZIP.
pause
exit /b 1
