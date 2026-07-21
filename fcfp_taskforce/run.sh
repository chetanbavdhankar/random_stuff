#!/bin/bash
# Move to script's directory so relative paths work regardless of launch location
cd "$(dirname "$0")"

echo "========================================================"
echo " Global FCFP AI Productivity — Use Case Hub Launcher"
echo "========================================================"
echo ""

# Find python command
if command -v python3 &>/dev/null; then
    PYTHON_CMD="python3"
elif command -v python &>/dev/null; then
    PYTHON_CMD="python"
else
    echo "ERROR: Python 3 is not installed or not in system PATH."
    echo "Please install Python 3.9+."
    read -p "Press Enter to exit..."
    exit 1
fi

# Ensure required packages are installed
$PYTHON_CMD -c "import flask, openpyxl" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "Installing required Python packages..."
    $PYTHON_CMD -m pip install -r requirements.txt
fi

echo "Starting local server at http://127.0.0.1:5000..."
echo "Your browser will open automatically in a moment."
echo "Keep this window open while using the application."
echo "========================================================"
echo ""

# Run Python server (browser auto-opens via server.py)
$PYTHON_CMD server.py
