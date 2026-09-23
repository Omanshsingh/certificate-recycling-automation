@echo off
setlocal
cd /d "%~dp0"
powershell.exe -NoProfile -File "%~dp0Install Daily Task.ps1" -Mode Setup
set "SETUP_RESULT=%errorlevel%"
pause
exit /b %SETUP_RESULT%
