@echo off
:: CCTV Bot Launcher — runs Python as Administrator
:: Double-click this file every morning to run the bot

:: Check if already running as admin
net session >nul 2>&1
if %errorLevel% == 0 (
    echo Running as Administrator...
    goto :run
) else (
    echo Requesting Administrator access...
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

:run
:: Change to the folder where this bat file is located
cd /d "%~dp0"

echo ==========================================
echo   CCTV Daily Check Bot v5
echo   %date% %time%
echo ==========================================
echo.

:: Check Python is installed
python --version >nul 2>&1
if %errorLevel% neq 0 (
    echo ERROR: Python not found!
    echo Please install Python from https://python.org
    echo Make sure to tick "Add Python to PATH" during install
    pause
    exit /b
)

:: Install requirements silently if needed
echo Checking requirements...
pip install -r requirements.txt -q --disable-pip-version-check

echo.
echo Starting camera check...
echo.

:: Run the bot
python cctv_bot.py

echo.
echo ==========================================
echo   Done! Check your Desktop for:
echo   - CCTV_Report_%date:~-4,4%-%date:~-7,2%-%date:~-10,2%.xlsx
echo   - CCTV_Snapshots folder
echo ==========================================
echo.
pause
