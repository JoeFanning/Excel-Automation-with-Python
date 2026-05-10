#!/bin/bash

# Move to the directory where this script is located
cd "$(dirname "$0")"

echo "---------------------------------------"
echo "Starting Excel Sales Automation (Mac/Linux)..."
echo "---------------------------------------"

# 1. Check for virtual environment; create if missing
if [ ! -d "venv" ]; then
    echo "Creating virtual environment..."
    python3 -m venv venv
fi

# 2. Activate and install dependencies from pyproject.toml
echo "Checking dependencies..."
source venv/bin/activate
pip install -e .

# 3. Run the main script
echo "Running Automation..."
python3 main.py

echo "---------------------------------------"
echo "Done! Press any key to close."
read -n 1

