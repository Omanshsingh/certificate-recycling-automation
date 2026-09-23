@echo off
setlocal
cd /d "%~dp0"
powershell.exe -NoProfile -File "%~dp0Install Daily Task.ps1" -Mode Install
if errorlevel 1 (
  echo The daily task could not be created. Follow the error above; ask IT if user tasks are blocked.
  pause
  exit /b 1
)
echo The daily check runs quietly at 09:00 and two minutes after sign-in.
pause
exit /b 0
