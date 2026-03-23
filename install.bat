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
echo   word_replace Setup
echo ======================================
echo.

:: 1. 检测环境
call check_env.bat

:: 2. 安装依赖
echo [2/3] 正在安装 Python 依赖...
python -m pip install --no-index --find-links=offline_libs python-docx lxml typing_extensions

:: 3. 创建桌面快捷方式
echo [3/3] 正在创建桌面快捷方式...
powershell -Command "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut('%USERPROFILE%\Desktop\word_replace-Word校对.lnk'); $s.TargetPath = '%CD%\启动Word校对.bat'; $s.WorkingDirectory = '%CD%'; $s.Save()"
powershell -Command "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut('%USERPROFILE%\Desktop\word_replace-知识库.lnk'); $s.TargetPath = '%CD%\启动知识库管理器.bat'; $s.WorkingDirectory = '%CD%'; $s.Save()"

echo [OK] Done!
echo.
pause
