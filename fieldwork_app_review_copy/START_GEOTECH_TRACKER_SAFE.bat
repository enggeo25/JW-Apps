@echo off
title Geotech Tracker Launcher
cd /d "%~dp0"

echo ==========================================
echo        Geotech Tracker - Starting
echo ==========================================
echo.
echo Leave this window open while using the app.
echo.

set "APP_PYTHON="
if exist "venv\Scripts\python.exe" set "APP_PYTHON=venv\Scripts\python.exe"

if not defined APP_PYTHON (
    where python >nul 2>nul
    if errorlevel 1 (
        echo Python was not found on this PC.
        echo Install Python or copy this app with its venv folder included.
        echo.
        pause
        exit /b
    )
    set "APP_PYTHON=python"
)

"%APP_PYTHON%" -c "import flask" >nul 2>nul
if errorlevel 1 (
    echo Installing required packages...
    "%APP_PYTHON%" -m pip install -r requirements.txt
    if errorlevel 1 (
        echo Failed to install requirements.
        echo.
        pause
        exit /b
    )
)

echo Opening app in your browser...
start "" http://127.0.0.1:5000

echo Starting server...
"%APP_PYTHON%" app.py

echo.
echo The server has stopped.
pause
