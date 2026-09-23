@echo off
setlocal
cd /d "%~dp0"
powershell.exe -NoProfile -File "%~dp0Install Daily Task.ps1" -Mode Desktop
set "APP_RESULT=%errorlevel%"
if not "%APP_RESULT%"=="0" pause
exit /b %APP_RESULT%
