#!/bin/bash
# Toyo Schedule Maker (Light) - Launcher
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
VENV_DIR="$HOME/.toyo_scheduler/venv"
VENV_PYTHON="$VENV_DIR/bin/python3"

# If venv or core deps missing, run the GUI installer
if [ ! -f "$VENV_PYTHON" ] || ! "$VENV_PYTHON" -c "import openpyxl" 2>/dev/null; then
    echo "Running first-time setup..."
    python3 "$SCRIPT_DIR/installer.py"
    # Check if install succeeded
    if [ ! -f "$VENV_PYTHON" ] || ! "$VENV_PYTHON" -c "import openpyxl" 2>/dev/null; then
        echo "Setup was cancelled or failed. Exiting."
        exit 1
    fi
fi

"$VENV_PYTHON" "$SCRIPT_DIR/toyo_scheduler_light.py"
