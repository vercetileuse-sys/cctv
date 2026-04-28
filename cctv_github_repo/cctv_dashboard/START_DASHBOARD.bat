@echo off
title CCTV Dashboard Server
color 0A

:: Auto-elevate to Administrator
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo Requesting Administrator access...
    powershell -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

cd /d "%~dp0"

echo ============================================
echo   CCTV Surveillance Dashboard
echo   Login: Admin / Auracctv#2024
echo ============================================
echo.
echo [1/3] Checking Python...
python --version >nul 2>&1
if %errorLevel% neq 0 (
    echo ERROR: Python not found!
    echo Install from https://python.org
    pause & exit /b
)
echo OK

echo [2/3] Installing requirements...
pip install flask requests openpyxl urllib3 pillow opencv-python -q --disable-pip-version-check
echo OK

echo [3/3] Starting server on port 5000...
echo.
echo  Dashboard: http://localhost:5000
echo  Network:   http://%COMPUTERNAME%:5000
echo.
echo  Press Ctrl+C to stop
echo ============================================
echo.

:: Open browser after 3 seconds
start "" /b cmd /c "timeout /t 3 /nobreak >nul && start http://localhost:5000"

python app.py

pause
