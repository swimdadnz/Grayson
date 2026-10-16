@echo off
setlocal
cd /d "%~dp0"
echo IMPORTANT: complete update\input\new_swimcloud_profiles_pending.csv with genuine SwimCloud URL + numeric ID for the added swimmers.
echo Ian Thorpe is prefilled and can be acquired separately. Nothing is installed.
echo You will have time to clear the browser verification and then press Enter.
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\swimcloud_extra_profiles_once.ps1"
if errorlevel 1 echo Acquisition stopped. See any logged diagnostics in update\scripts.
pause
