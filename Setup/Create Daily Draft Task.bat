@echo off
setlocal
cd /d "%~dp0"
powershell.exe -NoProfile -File "%~dp0..\App\Install Daily Task.ps1" -Mode Install
if errorlevel 1 (
  echo The daily task could not be created. Follow the error above; ask IT if user tasks are blocked.
  pause
  exit /b 1
)
echo The daily check uses the reminder time saved in Certificate Automation Settings (default 09:00).
pause
exit /b 0
