@echo off
setlocal
cd /d "%~dp0"
schtasks /Create /SC DAILY /TN "Certificate Recycling Drafts" /TR "\"%~dp0Run Daily Drafts.bat\"" /ST 09:00 /F
if errorlevel 1 (
  echo Could not create the scheduled task. Ask IT to allow a daily user task.
  pause
  exit /b 1
)
echo Daily draft task created for 09:00.
echo Make sure settings.local.json has been created by the desktop app first.
pause
