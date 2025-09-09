@echo off
REM Launch script for Journal Entry ID Creator GUI (Windows)

cd /d "%~dp0"

echo Starting Journal Entry ID Creator...

REM Check if Python is available
python --version >nul 2>&1
if %errorlevel% equ 0 (
    python launch_gui.py
    goto end
)

python3 --version >nul 2>&1
if %errorlevel% equ 0 (
    python3 launch_gui.py
    goto end
)

echo Error: Python is not installed or not in PATH
echo Please install Python 3.7 or later from https://python.org
pause
exit /b 1

:end
pause
