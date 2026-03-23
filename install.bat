@echo off
setlocal enabledelayedexpansion

for /f "delims=" %%i in ('"%~dp0find_python.bat"') do set PYTHON=%%i

if "%PYTHON%"=="ERROR" (
    echo [ERROR] Python not found
    echo Please edit config.bat to set PYTHON_PATH
    pause
    exit /b 1
)

cls
echo ========================================
echo   clawTest Setup
echo ========================================
echo.

"%PYTHON%" --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python cannot run
    pause
    exit /b 1
)

for /f "tokens=*" %%i in ('"%PYTHON%" --version') do set PYTHON_VERSION=%%i
echo [OK] Python: !PYTHON_VERSION!
echo.

echo ========================================
echo   Installing dependencies...
echo ========================================
echo.

"%PYTHON%" -c "import docx" >nul 2>&1
if %errorlevel% neq 0 (
    echo Installing python-docx...
    if exist "offline_libs\*.whl" (
        "%PYTHON%" -m pip install offline_libs\*.whl --no-index --find-links offline_libs
    ) else (
        "%PYTHON%" -m pip install python-docx
    )
    echo [OK] python-docx installed
) else (
    echo [OK] python-docx already installed
)
echo.

echo ========================================
echo   Creating shortcuts...
echo ========================================
echo.

powershell -Command "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut('%USERPROFILE%\Desktop\clawTest-Word校对.lnk'); $s.TargetPath = '%CD%\启动Word校对.bat'; $s.WorkingDirectory = '%CD%'; $s.Save()"

powershell -Command "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut('%USERPROFILE%\Desktop\clawTest-知识库.lnk'); $s.TargetPath = '%CD%\启动知识库管理器.bat'; $s.WorkingDirectory = '%CD%'; $s.Save()"

echo [OK] Done!
echo.
pause
