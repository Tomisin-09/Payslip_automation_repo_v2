@echo off
setlocal enabledelayedexpansion

echo === Payslip Automation Runner ===

REM ----------------------------------------
REM 1. Detect Python
REM ----------------------------------------
where python >nul 2>nul
if %ERRORLEVEL%==0 (
    set PY=python
) else (
    where py >nul 2>nul
    if %ERRORLEVEL%==0 (
        set PY=py
    ) else (
        echo.
        echo Python is not installed or not on PATH.
        echo.
        echo Please install Python 3.10 or newer from:
        echo https://www.python.org/downloads/windows/
        echo.
        echo IMPORTANT:
        echo - Tick "Add Python to PATH"
        echo - Restart your computer or Command Prompt
        pause
        exit /b 1
    )
)

%PY% --version

REM ----------------------------------------
REM 2. Create virtual environment
REM ----------------------------------------
if not exist .venv (
    echo Creating virtual environment...
    %PY% -m venv .venv
)

REM ----------------------------------------
REM 3. Activate virtual environment
REM ----------------------------------------
call .venv\Scripts\activate.bat

REM ----------------------------------------
REM 4. Install dependencies
REM ----------------------------------------
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

REM ----------------------------------------
REM 5. Run application
REM ----------------------------------------
python -m src.main

echo === Run complete ===
pause
