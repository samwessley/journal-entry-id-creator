#!/bin/bash
# Launch script for Journal Entry ID Creator GUI (macOS/Linux)

cd "$(dirname "$0")"

echo "Starting Journal Entry ID Creator..."

# Check if Python 3 is available
if command -v python3 &> /dev/null; then
    python3 launch_gui.py
elif command -v python &> /dev/null; then
    python launch_gui.py
else
    echo "Error: Python is not installed or not in PATH"
    echo "Please install Python 3.7 or later"
    read -p "Press Enter to exit..."
    exit 1
fi
