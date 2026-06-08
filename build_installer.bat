@echo off
title Building Student Result Analyzer Setup Wizard
cd /d "%~dp0"

echo ===================================================
echo   📦 Building Setup Wizard Executable (.exe)
echo ===================================================
echo.

:: 1. Check for virtual environment and activate
if exist venv (
    echo [INFO] Activating virtual environment...
    call venv\Scripts\activate
)

:: 2. Ensure PyInstaller is installed
echo [INFO] Ensuring PyInstaller is installed...
python -m pip install pyinstaller
if %errorlevel% neq 0 (
    echo [ERROR] Failed to install PyInstaller.
    pause
    exit /b 1
)

:: 3. Prepare compilation directories and copy assets for PyInstaller resolution
echo [INFO] Preparing build folders...
if not exist build_out mkdir build_out
if not exist build_out\build mkdir build_out\build
if not exist build_out\dist mkdir build_out\dist
if not exist build_out\assets mkdir build_out\assets

echo [INFO] Copying icon for spec resolution...
copy /y assets\icon.ico build_out\assets\icon.ico >nul

:: 4. Automatically close any running instances of the installer to avoid file lock
echo [INFO] Closing any running Setup instances...
taskkill /f /im Setup_StudentResultAnalyzer.exe >nul 2>&1

:: 5. Compile installer.py to Setup_StudentResultAnalyzer.exe
echo.
echo [INFO] Compiling installer.py into Setup Wizard...
echo [INFO] This might take a minute, please wait...
pyinstaller --clean --noconfirm --onefile --windowed --icon=assets/icon.ico --name="Setup_StudentResultAnalyzer" --specpath="build_out" --workpath="build_out/build" --distpath="build_out/dist" installer.py

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] PyInstaller failed to build the Setup executable.
    pause
    exit /b 1
)

echo.
echo ===================================================
echo   🎉 Success! Setup Wizard created successfully!
echo   File: build_out\dist\Setup_StudentResultAnalyzer.exe
echo ===================================================
echo.
pause
