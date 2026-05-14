#!/bin/bash

# 1. Force the terminal to jump to the directory containing this script file
cd "$(dirname "$0")"

echo "==========================================="
echo "  EXCEL AUTOMATION SYSTEM - MAC STARTING"
echo "==========================================="

# 2. Setup Virtual Environment if missing
if [ ! -d "venv" ]; then
    echo "[1/2] Creating local virtual environment..."
    python3 -m venv venv

    echo "Updating internal deployment tools..."
    ./venv/bin/python -m pip install --upgrade pip -q

    echo "[2/2] Installing package dependencies..."
    ./venv/bin/python -m pip install -r requirements.txt -q
else
    echo "[1/2] Environment verified."
fi

# 3. Expose root path to Python so it maps the /src directory correctly
export PYTHONPATH=$(pwd)

# 4. Run the desktop launcher application
echo "[2/2] Launching Desktop GUI Pipeline..."
./venv/bin/python main.py

if [ $? -ne 0 ]; then
    echo "==========================================="
    echo "  ERROR: The application crashed or failed."
    echo "==========================================="
    read -p "Press enter to close..."
    exit 1
fi

echo "==========================================="
echo "  SUCCESS: Operations completed."
echo "==========================================="
