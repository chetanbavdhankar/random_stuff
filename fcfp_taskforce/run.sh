#!/bin/bash
# Move to script's directory so relative paths work regardless of launch location
cd "$(dirname "$0")"

echo "========================================================"
echo " Global FCFP AI Productivity — Use Case Hub Launcher"
echo "========================================================"
echo ""

PYTHON_CMD=""

# 1. Check 'py -3' (Official Python Launcher on Windows)
if command -v py &>/dev/null && py -3 -c "import sys" &>/dev/null; then
    PYTHON_CMD="py -3"
# 2. Check 'python3' (macOS / Linux / Windows)
elif command -v python3 &>/dev/null && python3 -c "import sys" &>/dev/null; then
    PYTHON_CMD="python3"
# 3. Check 'python' (macOS / Linux / Windows)
elif command -v python &>/dev/null && python -c "import sys" &>/dev/null; then
    PYTHON_CMD="python"
else
    # 4. Fallback search for common Windows installation locations in Git Bash
    for win_py in "$LOCALAPPDATA"/Programs/Python/Python3*/python.exe \
                  /c/Python3*/python.exe \
                  "/c/Program Files/Python3*/python.exe" \
                  "/c/Program Files (x86)/Python3*/python.exe"; do
        if [ -f "$win_py" ]; then
            PYTHON_CMD="$win_py"
            break
        fi
    done
fi

if [ -z "$PYTHON_CMD" ]; then
    echo "ERROR: Python 3 was not found on your system."
    echo "Please install Python 3 from https://www.python.org/ or Microsoft Store."
    echo "If Python is already installed, make sure 'Add python.exe to PATH' was checked during installation."
    read -p "Press Enter to exit..."
    exit 1
fi

run_py() {
    if [ "$PYTHON_CMD" = "py -3" ]; then
        py -3 "$@"
    else
        "$PYTHON_CMD" "$@"
    fi
}

echo "Using Python ($PYTHON_CMD)..."

# Ensure required packages are installed
run_py -c "import flask, openpyxl" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "Installing required Python packages (flask, openpyxl)..."
    run_py -m pip install -r requirements.txt
fi

echo "Starting local server at http://127.0.0.1:5000..."
echo "Your browser will open automatically in a moment."
echo "Keep this window open while using the application."
echo "========================================================"
echo ""

# Run Python server (browser auto-opens via server.py)
run_py server.py
