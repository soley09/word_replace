@echo off
REM Auto-find Python for word_replace

REM Check Miniconda first (most common on this machine)
if exist "D:\Miniconda\python.exe" (
    echo D:\Miniconda\python.exe
    exit /b 0
)

REM Check Miniconda3
if exist "D:\Miniconda3\python.exe" (
    echo D:\Miniconda3\python.exe
    exit /b 0
)

REM Check user config
if exist "%~dp0config.bat" (
    call "%~dp0config.bat"
    if defined PYTHON_PATH (
        if exist "%PYTHON_PATH%" (
            echo %PYTHON_PATH%
            exit /b 0
        )
    )
)

REM Check system PATH
where python >nul 2>&1
if %errorlevel% equ 0 (
    for /f "delims=" %%i in ('where python') do (
        REM Skip Windows Store stub
        echo %%i | findstr /i "WindowsApps" >nul
        if errorlevel 1 (
            echo %%i
            exit /b 0
        )
    )
)

echo ERROR
exit /b 1
