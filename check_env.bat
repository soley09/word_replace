@echo off
chcp 65001 >nul

cls
echo ========================================
echo   clawTest 环境检测
echo ========================================
echo.

REM 检查Python
echo [Python]
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo   ❌ 未安装
    echo   下载地址: https://www.python.org/downloads/
) else (
    for /f "tokens=*" %%i in ('python --version') do echo   ✅ %%i
)
echo.

REM 检查pip
echo [pip]
pip --version >nul 2>&1
if %errorlevel% neq 0 (
    echo   ❌ 未安装
) else (
    for /f "tokens=*" %%i in ('pip --version') do echo   ✅ %%i
)
echo.

REM 检查python-docx
echo [python-docx]
python -c "import docx" >nul 2>&1
if %errorlevel% neq 0 (
    echo   ❌ 未安装
    echo   运行: pip install python-docx
) else (
    echo   ✅ 已安装
)
echo.

REM 检查pywin32
echo [pywin32 (可选)]
python -c "import win32com" >nul 2>&1
if %errorlevel% neq 0 (
    echo   ⚠️ 未安装 (可选)
) else (
    echo   ✅ 已安装
)
echo.

echo ========================================
echo   检测完成
echo ========================================
pause
