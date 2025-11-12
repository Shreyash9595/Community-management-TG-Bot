@echo off
title Hunt Report Generator Bot
color 0A

echo =====================================================
echo ✅ Launching Hunt Report Generator Bot...
echo =====================================================

REM Activate virtual environment
call "%~dp0venv\Scripts\activate"

REM Run the bot
python "%~dp0main.py"

REM Optional: pause to see any error messages before closing
echo -----------------------------------------------------
echo Bot has stopped running. Press any key to exit.
pause >nul
