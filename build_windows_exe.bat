@echo off
setlocal enableextensions enabledelayedexpansion

REM Build MailNotifier.exe on Windows using PyInstaller.
REM This script builds from app_tk.py (recommended for Windows).

cd /d "%~dp0"

set "LOG=%cd%\build_windows_exe.log"
echo. > "%LOG%"
echo [build] started at %date% %time%>> "%LOG%"

if not exist "app_tk.py" (
  echo ERROR: app_tk.py not found in %cd%
  echo Run this .bat from the project folder.
  echo ERROR: app_tk.py not found>> "%LOG%"
  pause
  exit /b 1
)

REM Pick Python launcher if available, otherwise fallback to python.
set "PY="
where py >nul 2>&1
if not errorlevel 1 (
  set "PY=py -3"
)
if "%PY%"=="" (
  where python >nul 2>&1
  if not errorlevel 1 (
    set "PY=python"
  )
)

if "%PY%"=="" (
  echo ERROR: Python not found.
  echo Install Python 3.10+ from python.org and re-run this file.
  echo Tip: disable Windows "App execution aliases" for python.exe if needed.
  echo ERROR: Python not found>> "%LOG%"
  pause
  exit /b 1
)

echo Using: %PY%
echo [build] Using: %PY%>> "%LOG%"

echo Checking Python...
%PY% --version >> "%LOG%" 2>&1
%PY% --version
if errorlevel 1 (
  echo ERROR: Python command failed. See %LOG%
  echo ERROR: python --version failed>> "%LOG%"
  pause
  exit /b 1
)

REM Create venv if missing
if not exist ".venv\Scripts\python.exe" (
  echo Creating venv...
  %PY% -m venv .venv >> "%LOG%" 2>&1
  if errorlevel 1 goto :fail
)

call ".venv\Scripts\activate.bat"
if errorlevel 1 goto :fail

echo Installing dependencies...
python -m pip install --upgrade pip >> "%LOG%" 2>&1
if errorlevel 1 goto :fail
python -m pip install -r requirements.txt >> "%LOG%" 2>&1
if errorlevel 1 goto :fail
python -m pip install pyinstaller >> "%LOG%" 2>&1
if errorlevel 1 goto :fail

echo Building exe...
pyinstaller --noconfirm --clean --windowed --name "MailNotifier" app_tk.py >> "%LOG%" 2>&1
if errorlevel 1 goto :fail

echo.
echo DONE.
echo Your exe is here:
echo   %cd%\dist\MailNotifier\MailNotifier.exe
echo [build] DONE>> "%LOG%"
echo.
pause
exit /b 0

:fail
echo.
echo BUILD FAILED.
echo See log file:
echo   %LOG%
echo [build] FAILED>> "%LOG%"
pause
exit /b 1

