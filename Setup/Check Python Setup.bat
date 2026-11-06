@echo off
setlocal
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0..\App\Install Daily Task.ps1" -Mode Diagnose
echo.
pause
