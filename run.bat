@echo off
title Student Result Analyzer - VNSGU
setlocal enabledelayedexpansion

echo ===================================================
echo   🎓 Student Result Analyzer - SASCMA ^| VNSGU
echo ===================================================
echo.

:: Check if Python is installed
where python >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python is not installed or not in your PATH.
    echo Please download and install Python 3.10+ from:
    echo https://www.python.org/downloads/
    echo.
    echo IMPORTANT: Make sure to check the checkbox that says:
    echo "Add Python.exe to PATH" during installation.
    echo.
    pause
    exit /b 1
)

:: Check if we are in the correct directory (directory of the batch file)
cd /d "%~dp0"

:: Check if requirements.txt exists
if not exist requirements.txt (
    echo [ERROR] requirements.txt not found in the current directory.
    echo Please run this script from the project folder containing 'main.py' and 'requirements.txt'.
    echo.
    pause
    exit /b 1
)

:: Check for virtual environment
if not exist venv (
    echo [INFO] Virtual environment 'venv' not found. Creating a new one...
    python -m venv venv
    if !errorlevel! neq 0 (
        echo [ERROR] Failed to create virtual environment.
        pause
        exit /b 1
    )
    echo [INFO] Virtual environment created successfully.
    echo.
    echo [INFO] Activating virtual environment and installing dependencies...
    call venv\Scripts\activate
    python -m pip install --upgrade pip
    pip install -r requirements.txt
    if !errorlevel! neq 0 (
        echo [ERROR] Failed to install dependencies.
        pause
        exit /b 1
    )
    echo [INFO] Dependencies installed successfully.
) else (
    echo [INFO] Activating virtual environment...
    call venv\Scripts\activate
)

echo.
echo [INFO] Starting Student Result Analyzer...
python main.py
if %errorlevel% neq 0 (
    echo.
    echo [WARNING] Application exited with an error code: %errorlevel%
    pause
)
