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

    :: 5. Install libraries
    echo [2/2] Installing libraries...
    pip install . --use-feature=in-tree-build -q
) else (
    :: 6. Just activate
    echo [1/2] Environment ready.
    call venv\Scripts\activate
)

:: 7. Run the Python automation script
echo [2/2] Processing Excel files...
python main.py

:: 8. Final success message
echo ===========================================
echo   SUCCESS: Reports generated in /output
echo ===========================================
pause
