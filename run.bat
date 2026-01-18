@echo off
setlocal EnableExtensions EnableDelayedExpansion
cd /d "%~dp0"

REM ==============================
REM 1) Find Python (prefer py, fallback to python)
REM ==============================
set "PYCMD="

where py >nul 2>&1
if %errorlevel%==0 set "PYCMD=py"

if not defined PYCMD (
  where python >nul 2>&1
  if %errorlevel%==0 set "PYCMD=python"
)

if not defined PYCMD (
  echo.
  echo ERROR: Python 3.10+ is required but was not found.
  echo Install from https://www.python.org (tick "Add python.exe to PATH").
  echo.
  pause
  exit /b 1
)

REM ==============================
REM 2) Enforce minimum version (>= 3.10)
REM ==============================
for /f "usebackq delims=" %%V in (`%PYCMD% -c "import sys; print(f'{sys.version_info[0]}.{sys.version_info[1]}')" 2^>nul`) do set "PYVER=%%V"

if not defined PYVER (
  echo.
  echo ERROR: Python was found (%PYCMD%) but version could not be detected.
  echo Try running: %PYCMD% --version
  echo.
  pause
  exit /b 1
)

for /f "tokens=1,2 delims=." %%a in ("%PYVER%") do (
  set "MAJOR=%%a"
  set "MINOR=%%b"
)

if %MAJOR% LSS 3 goto :badver
if %MAJOR% EQU 3 if %MINOR% LSS 10 goto :badver

echo Python version OK: %PYVER% (via %PYCMD%)

REM ==============================
REM 3) venv + deps + run
REM ==============================
if not exist ".venv\Scripts\python.exe" (
  echo Creating virtual environment...
  %PYCMD% -m venv .venv
  if errorlevel 1 (
    echo ERROR: Failed to create virtual environment.
    pause
    exit /b 1
  )
)

echo Upgrading pip...
".venv\Scripts\python.exe" -m pip install --upgrade pip
if errorlevel 1 (
  echo ERROR: Failed to upgrade pip.
  pause
  exit /b 1
)

echo Installing requirements...
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 (
  echo ERROR: Failed to install requirements.
  pause
  exit /b 1
)

echo Running...
".venv\Scripts\python.exe" -m src.main
set "EXITCODE=%ERRORLEVEL%"

echo.
if not "%EXITCODE%"=="0" (
  echo Process finished with errors. Exit code: %EXITCODE%
) else (
  echo Done.
)
pause
exit /b %EXITCODE%

:badver
echo.
echo ERROR: Python 3.10+ is required. Found: %PYVER%
echo Please install a newer version from https://www.python.org
echo.
pause
exit /b 1