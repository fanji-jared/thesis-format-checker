@echo off
chcp 65001 >nul
title 论文格式检测工具

echo ========================================
echo    论文格式检测工具 启动中...
echo ========================================
echo.

cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo [错误] 未找到虚拟环境，请先创建虚拟环境
    echo.
    echo 创建虚拟环境命令:
    echo   python -m venv .venv
    echo   .venv\Scripts\activate
    echo   pip install -r requirements.txt
    echo.
    pause
    exit /b 1
)

echo [信息] 正在启动程序...
echo.

.venv\Scripts\python.exe main.py %*

if errorlevel 1 (
    echo.
    echo [错误] 程序运行出错
    pause
)
