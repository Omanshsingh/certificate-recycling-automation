@echo off
setlocal
cd /d "%~dp0"
set "PYTHON_CMD="
where python >nul 2>nul
if not errorlevel 1 (
  python -c "import sys; raise SystemExit(0 if sys.version_info >= (3, 11) else 1)" >nul 2>nul
  if not errorlevel 1 set "PYTHON_CMD=python"
)
if not defined PYTHON_CMD (
  where py >nul 2>nul
  if not errorlevel 1 set "PYTHON_CMD=py -3"
)
if not defined PYTHON_CMD goto no_python
%PYTHON_CMD% -c "import sys; print('Found Python ' + sys.version.split()[0]); raise SystemExit(0 if sys.version_info >= (3, 11) else 1)"
if errorlevel 1 goto old_python
echo Installing required packages for this Windows user...
%PYTHON_CMD% -m pip install --user -r requirements.txt
if errorlevel 1 goto packages_failed
echo Setup complete. Double-click Prepare Certificates.bat.
pause
exit /b 0
:no_python
echo Python 3.11 or newer was not found. Install it from https://www.python.org/downloads/windows/
pause
exit /b 1
:old_python
echo The detected Python is older than 3.11. Install a newer version and try again.
pause
exit /b 1
:packages_failed
echo Required packages could not be installed. Check internet access or ask IT for help.
pause
exit /b 1
