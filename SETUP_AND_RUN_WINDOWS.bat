@echo off
setlocal
cd /d "%~dp0"

set "PY_CMD="
set "VENV_PY=%CD%\.venv\Scripts\python.exe"

call :detect_python
if not defined PY_CMD (
    echo Python was not found on this system.
    where winget >nul 2>&1
    if %errorlevel%==0 (
        echo Attempting to install Python 3.12 using winget...
        winget install -e --id Python.Python.3.12 --accept-package-agreements --accept-source-agreements
    ) else (
        echo winget is not available on this PC.
    )

    call :detect_python
    if not defined PY_CMD (
        echo ERROR: Python is still not available.
        echo Please install Python 3.10 or newer manually from https://www.python.org/downloads/
        echo Then re-run this script.
        pause
        exit /b 1
    )
)

%PY_CMD% -c "import sys; raise SystemExit(0 if sys.version_info >= (3, 10) else 1)"
if errorlevel 1 (
    echo ERROR: Python 3.10 or newer is required.
    %PY_CMD% -c "import sys; print('Detected Python version:', sys.version.split()[0])"
    pause
    exit /b 1
)

if not exist "%CD%\requirements.txt" (
    echo ERROR: requirements.txt was not found in the repository root.
    pause
    exit /b 1
)

if not exist "%CD%\.venv\Scripts\python.exe" (
    echo Creating virtual environment at .venv...
    %PY_CMD% -m venv .venv
    if errorlevel 1 (
        echo ERROR: Failed to create virtual environment.
        pause
        exit /b 1
    )
)

"%VENV_PY%" -m pip install --upgrade pip
if errorlevel 1 (
    echo ERROR: Failed to upgrade pip in .venv.
    pause
    exit /b 1
)

"%VENV_PY%" -m pip install -r requirements.txt
if errorlevel 1 (
    echo ERROR: Failed to install dependencies from requirements.txt.
    pause
    exit /b 1
)

"%VENV_PY%" main.py
if errorlevel 1 (
    echo ERROR: main.py exited with an error.
    pause
    exit /b 1
)

exit /b 0

:detect_python
set "PY_CMD="
where py >nul 2>&1
if %errorlevel%==0 (
    py -3 -c "import sys" >nul 2>&1
    if %errorlevel%==0 (
        set "PY_CMD=py -3"
        goto :eof
    )
)
where python >nul 2>&1
if %errorlevel%==0 (
    set "PY_CMD=python"
)
goto :eof
