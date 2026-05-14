@echo off
setlocal
cd /d "%~dp0"

echo ===========================================
echo   EXCEL AUTOMATION SYSTEM - STARTING
echo ===========================================

:: 1. Check if the virtual environment exists
if not exist venv (
    echo [1/2] First-time setup... creating environment.

    :: 2. Create environment
    python -m venv venv

    :: 3. Activate environment
    call venv\Scripts\activate

    :: 4. Update pip tool
    echo Updating internal tools...
    python -m pip install --upgrade pip -q

    :: 5. Install libraries like Pandas and Matplotlib
    echo [2/2] Installing libraries...
    pip install . -q

) else (
    :: 6. Just activate
    echo [1/2] Environment ready.
    cal@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo ===========================================
echo   EXCEL AUTOMATION SYSTEM - STARTING
echo ===========================================

:: 1. Verify Python is installed
where python >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Python is not installed or not in system PATH.
    pause
    exit /b 1
)

:: 2. Setup Virtual Environment if missing
if not exist venv (
    echo [1/2] Creating local virtual environment...
    python -m venv venv
    if %ERRORLEVEL% NEQ 0 (
        echo ERROR: Failed to create virtual environment.
        pause
        exit /b %ERRORLEVEL%
    )
    echo Updating internal deployment tools...
    "%~dp0venv\Scripts\python.exe" -m pip install --upgrade pip -q

    echo [2/2] Installing package dependencies...
    "%~dp0venv\Scripts\pip.exe" install -r requirements.txt -q
) else (
    echo [1/2] Environment verified.
)

:: 3. Set Python Path to find the /src directory modules properly
set PYTHONPATH=%~dp0

:: 4. Run the launcher
echo [2/2] Launching Desktop GUI Pipeline...
"%~dp0venv\Scripts\python.exe" main.py
if %ERRORLEVEL% NEQ 0 (
    echo ===========================================
    echo   ERROR: The application crashed or failed.
    echo ===========================================
    pause
    exit /b %ERRORLEVEL%
)

echo ===========================================
echo   SUCCESS: Operations completed.
echo ===========================================
pause
l venv\Scripts\activate
)

:: 7. Run the Python automation script
echo [2/2] Processing Excel files...
python main.py
if %ERRORLEVEL% NEQ 0 (
    echo ===========================================
    echo   ERROR: The script failed to run.
    echo ===========================================
    pause
    exit /b %ERRORLEVEL%
)

:: 8. Final success message
echo ===========================================
echo   SUCCESS: Reports generated in /output
echo ===========================================
pause
