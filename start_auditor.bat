@echo off
title AI based Intelligent PFMEA Auditor
echo ====================================================
echo Starting AI based Intelligent PFMEA Auditor Server...
echo ====================================================

:: Navigate to script directory
cd /d "%~dp0"

:: Prefer the project virtual environment when it exists.
if exist ".venv\Scripts\python.exe" (
    set "PYTHON=.venv\Scripts\python.exe"
) else (
    set "PYTHON=python"
)

if "%PYTHON%"=="python" (
    where python >nul 2>&1
    if errorlevel 1 (
        echo [ERROR] Python was not found on PATH. Install Python or create .venv first.
        pause
        exit /b 1
    )
) else if not exist "%PYTHON%" (
    echo [ERROR] Python was not found. Install Python or create .venv first.
    pause
    exit /b 1
)

:: Check if .env file exists
if not exist ".env" (
    echo [WARNING] .env file not found in directory!
    echo Please make sure .env contains API_KEY_S30 or API_KEY_Pro.
    echo.
)

:: Wait 2 seconds and open default browser to the web interface
start "" http://127.0.0.1:8000

:: Run the FastAPI Uvicorn server
"%PYTHON%" -m uvicorn app:app --host 127.0.0.1 --port 8000

pause
