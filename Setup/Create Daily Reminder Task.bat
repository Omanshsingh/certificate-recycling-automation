@echo off
setlocal
cd /d "%~dp0.."
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0..\App\Install Daily Task.ps1" -Mode Install
if errorlevel 1 (
  echo The daily reminder could not be enabled. Correct the message above and try again.
  pause
  exit /b 1
)
echo The daily alert checks the two Excel files and does not generate PDFs or send mail.
pause
exit /b 0
