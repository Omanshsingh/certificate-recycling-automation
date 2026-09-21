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
if not defined PYTHON_CMD (
  echo Python was not found. Run Setup Once.bat first.
  pause
  exit /b 1
)
%PYTHON_CMD% -c "import openpyxl, reportlab, win32com.client" >nul 2>nul
if errorlevel 1 (
  echo Installing missing certificate automation packages...
  %PYTHON_CMD% -m pip install --user -r requirements.txt
  if errorlevel 1 (
    echo Package installation failed. Check internet access or ask IT for help.
    pause
    exit /b 1
  )
)
%PYTHON_CMD% desktop_app.py
if errorlevel 1 pause
