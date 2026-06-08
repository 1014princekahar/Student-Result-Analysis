@echo off
title Building Student Result Analyzer EXE
cd /d "%~dp0"

echo ===================================================
echo   📦 Building Student Result Analyzer EXE & Setup
echo ===================================================
echo.

:: Check for virtual environment
if exist venv (
    echo [INFO] Activating virtual environment...
    call venv\Scripts\activate
)

echo [INFO] Installing PyInstaller...
python -m pip install pyinstaller
if %errorlevel% neq 0 (
    echo [ERROR] Failed to install PyInstaller.
    pause
    exit /b 1
)

:: Create build output directory
if not exist build_out mkdir build_out

:: Automatically close any running instances of the app to avoid WinError 5 (Access Denied)
echo.
echo [INFO] Closing any running instances of the application...
taskkill /f /im StudentResultAnalyzer.exe >nul 2>&1
taskkill /f /im Setup_StudentResultAnalyzer.exe >nul 2>&1

echo.
echo [INFO] Choose what you want to build:
echo [1] Build Standalone App (StudentResultAnalyzer.exe)
echo [2] Build Setup Wizard (Setup_StudentResultAnalyzer.exe)
echo [3] Build Both (Default)
echo.
set /p choice="Enter your choice (1/2/3): "

if "%choice%"=="1" goto :build_app
if "%choice%"=="2" goto :build_setup
if "%choice%"=="3" goto :build_both
goto :build_both

:build_app
echo.
echo [INFO] Compiling main.py into standalone App...
pyinstaller --noconfirm --onefile --windowed --add-data "assets;assets" --add-data "core;core" --icon="%~dp0assets\icon.ico" --name="StudentResultAnalyzer" --specpath="build_out" --workpath="build_out/build" --distpath="build_out/dist" main.py
goto :build_finish

:build_setup
echo.
echo [INFO] Compiling installer.py into Setup Wizard...
pyinstaller --noconfirm --onefile --windowed --icon="%~dp0assets\icon.ico" --name="Setup_StudentResultAnalyzer" --specpath="build_out" --workpath="build_out/build" --distpath="build_out/dist" installer.py
goto :build_finish

:build_both
echo.
echo [INFO] Compiling main.py into standalone App...
pyinstaller --noconfirm --onefile --windowed --add-data "assets;assets" --add-data "core;core" --icon="%~dp0assets\icon.ico" --name="StudentResultAnalyzer" --specpath="build_out" --workpath="build_out/build" --distpath="build_out/dist" main.py

echo.
echo [INFO] Compiling installer.py into Setup Wizard...
pyinstaller --noconfirm --onefile --windowed --icon="%~dp0assets\icon.ico" --name="Setup_StudentResultAnalyzer" --specpath="build_out" --workpath="build_out/build" --distpath="build_out/dist" installer.py
goto :build_finish

:build_finish
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] PyInstaller failed to build the executable(s).
    pause
    exit /b 1
)

echo.
echo ===================================================
echo   🎉 Success! Executable(s) created in the 'build_out/dist' folder.
echo   Path: build_out\dist\
echo ===================================================
echo.
pause
