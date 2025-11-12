@echo off
title 🧩 Hunt Report Generator Setup Wizard
chcp 65001 >nul
color 0A

echo =====================================================
echo      🧩 Hunt Report Generator Setup Wizard
echo =====================================================

:: Step 1: Check Python
echo.
echo [1/5] Checking for Python installation...
python --version >nul 2>&1
if errorlevel 1 (
    echo ⚠️ Python not found. Installing Python 3.11.9...
    powershell -Command "Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.11.9/python-3.11.9-amd64.exe' -OutFile 'python_installer.exe'"
    start /wait python_installer.exe /quiet InstallAllUsers=1 PrependPath=1 Include_test=0
    del python_installer.exe
) else (
    echo ✅ Python is already installed.
)

:: Step 2: Create venv if not exists
echo.
echo [2/5] Setting up virtual environment...
if not exist venv (
    python -m venv venv
    echo ✅ Virtual environment created.
) else (
    echo Virtual environment already exists.
)

:: Step 3: Activate venv
echo.
echo [3/5] Activating environment...
call venv\Scripts\activate
echo Environment activated.

:: Step 4: Check stdlib (imghdr fix)
echo.
echo [4/5] Checking Python standard library integrity...
python -c "import imghdr" 2>nul
if errorlevel 1 (
    echo ⚠️  Missing built-in modules detected. Rebuilding environment...
    rmdir /s /q venv
    python -m venv venv
    call venv\Scripts\activate
    echo 🧱 Environment rebuilt successfully.
) else (
    echo ✅ Python standard library verified.
)

:: Step 5: Install dependencies
echo.
echo [5/5] Installing dependencies...
pip install --upgrade pip >nul
pip install python-telegram-bot==13.15 pandas openpyxl pillow >nul
echo ✅ Dependencies installed.

:: Step 6: Verify folders
echo.
echo [6/6] Checking folder structure...
for %%F in (downloads uploads processed outputs) do (
    if not exist %%F (
        mkdir %%F
        echo Created fold
