#!/bin/bash
# ============================================================
#   HLR Checker - launcher (macOS / Linux)
#   Analog of start.bat for Windows
#   Run: double-click in Finder (after: chmod +x start.command)
# ============================================================

# Go to the folder where this script is located
cd "$(dirname "$0")" || exit 1

echo "============================================================"
echo "  HLR Checker"
echo "============================================================"
echo ""

# Find Python 3
PYTHON=""
if command -v python3 >/dev/null 2>&1; then
    PYTHON="python3"
elif command -v python >/dev/null 2>&1; then
    PYTHON="python"
else
    echo "[ERROR] Python 3 not found."
    echo "Install it from python.org or via: brew install python"
    echo ""
    read -r -p "Press Enter to exit..."
    exit 1
fi

echo "[1/2] Installing dependencies..."
"$PYTHON" -m pip install -r requirements.txt --quiet

echo "[2/2] Running..."
echo ""
"$PYTHON" main.py

echo ""
read -r -p "Press Enter to exit..."
