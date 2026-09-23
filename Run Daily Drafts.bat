@echo off
setlocal
cd /d "%~dp0"
powershell.exe -NoProfile -File "%~dp0Install Daily Task.ps1" -Mode Daily
exit /b %errorlevel%
