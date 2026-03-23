@echo off
cd /d "%~dp0"

for /f "delims=" %%i in ('"%~dp0find_python.bat"') do set PYTHON=%%i

if "%PYTHON%"=="ERROR" (
    echo [ERROR] Python not found
    echo Please edit config.bat to set PYTHON_PATH
    pause
    exit /b 1
)

"%PYTHON%" run.py
pause
